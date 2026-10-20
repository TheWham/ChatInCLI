# -*- coding: utf-8 -*-
"""微信命令行摸鱼聊天工具（适配微信 4.x · 全程无感后台版）

特性与关键修复：
1. 严格槽位对齐 + 按名坐标缓存：无效坐标用 None 占位保住行号对齐，杜绝下标错位点错人。
2. 状态保持（stealth_window_context）：原本最小化才在采集后最小化回去，不强行最小化已展开窗口；
   但**无论如何进入前都先把窗口置成全透明**，否则微信开着时，库里 open_weixin 改尺寸、
   搬位置的动作会被直接看见。
3. 精准过滤公众号/视频号：仅匹配系统级标签与公众号名，避免好友消息含「视频号」字样被误杀。
4. 群聊屏蔽：会话列表的 UIA 里群聊与单聊结构完全一致，认不出，所以走三层
   —— 名字含「群/组」+ 手动名单 + 打开过的群由 is_group_chat 认出后永久记住。
5. 后台重连熔断：连续失败 3 次即暂停自动重连，防止每 1.5 秒恢复/最小化风暴；
   熔断期间其他好友的新消息提醒照常。
6. 窗口副作用隔离（preserved_window_geometry）：pyweixin 的 Navigator.open_weixin 会把
   微信窗口改成 1000x1000 居中并设成置顶且从不复原，调用前后把这三点还原。
7. 焦点保护（preserved_foreground）：set_focus() 把前台抢到微信时自动还给终端。
8. 历史消息回看：切换好友默认展示最近 5 条记录，支持 /h [N] 随时回看。
9. 会话列表：启动与 /list 走零打扰的可见读取（约 9 条，稳定）；/list all 会滚动读取
   最多 30 条，但 Messages.dump_sessions 实测不稳定（连续两次调用会 41/41/0/0），
   失败时自动退回可见读取并说明原因。
10. /clear（同 /cls）清屏：只清终端画面上的内容，聊天记录一条不动。
"""
import contextlib
import json
import os
import threading
import time
from pathlib import Path
import win32api
import win32gui
import win32con
import win32process
from pywinauto import Desktop

from pyweixin import Navigator, Messages, GlobalConfig, Tools
from pyweixin.Uielements import (Lists, Edits, Buttons, Texts, Special_Labels,
                                 Main_window, Regex_Patterns)


# ---------------------------------------------------------------------------
# 过滤与历史记录配置
# ---------------------------------------------------------------------------
DEFAULT_HISTORY_COUNT = 5
MAX_RECONNECT_FAILURES = 3

# 仅对会话名称（name）做精确匹配的屏蔽集合
IGNORED_SESSION_NAMES = {
    '公众号', '订阅号', '订阅号消息', '服务号', '服务通知',
    '视频号', '视频号消息', '微信支付', '微信运动', '微信团队',
    '腾讯新闻', '微信游戏', '小程序', '折叠的群聊', '折叠置顶聊天',
    '企业微信通知',
}

# 仅匹配微信系统级带方括号的预览前缀，避免好友普通聊天提到「视频号」被误杀
IGNORED_PREVIEW_TAGS = (
    '[视频号]', '[服务通知]', '[公众号]', '[订阅号]',
)

# 聊天窗口内不显示的消息类型（对应 classify 返回的类型名）
IGNORED_MSG_TYPES = {'视频号'}


# ---------------------------------------------------------------------------
# 群聊屏蔽
# ---------------------------------------------------------------------------
# 实测：微信会话列表的 UIA 里，群聊和单聊没有任何可判别的结构差异 ——
# class_name（都是 mmui::ChatSessionCell）、control_type、automation_id、
# 子元素列表全都一模一样，所以只能靠名字特征 + 手动名单 + 「打开过的自动记住」。
#
# 曾经试过用「预览形如 [N条] 发送人: 内容」来认群聊，实测会误伤单聊
# （例如好友发来「时间: 明天下午三点」，就长得一模一样），已弃用。
IGNORE_GROUP_CHATS = True

# 第 1 层：名字里含这些关键词就当群聊。
# '组' 会误伤名字带「组」的单聊（比如「组长」），觉得太宽就删掉它。
GROUP_NAME_KEYWORDS = ('群', '组')

# 第 2 层：手动名单。命名特征认不出来的群聊（名字不含「群/组」）可以写在这里。
# 也可以什么都不写 —— 切进去一次后会被 is_group_chat 认出来，自动记进
# blocked_groups.json（该文件已 gitignore，不会进仓库）。
IGNORED_GROUP_NAMES = set()

# 第 3 层：打开过的群聊会被 is_group_chat 认出来并记到这里，下次启动仍生效。
BLOCKED_GROUPS_FILE = Path(__file__).with_name('blocked_groups.json')


def load_blocked_groups():
    """读回历史识别出的群聊（文件不存在或坏了都当作空集）"""
    try:
        with open(BLOCKED_GROUPS_FILE, encoding='utf-8') as f:
            names = json.load(f)
        if isinstance(names, list):
            IGNORED_GROUP_NAMES.update(n for n in names if isinstance(n, str) and n)
    except Exception:
        pass


def save_blocked_groups():
    try:
        with open(BLOCKED_GROUPS_FILE, 'w', encoding='utf-8') as f:
            json.dump(sorted(IGNORED_GROUP_NAMES), f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def is_group_session(name):
    """判断某个会话是不是群聊。只靠名字特征和历史名单，绝不猜消息内容"""
    return name in IGNORED_GROUP_NAMES or any(kw in name for kw in GROUP_NAME_KEYWORDS)


def is_ignored_session(name, raw_text=''):
    """判断某个会话是否属于公众号/视频号/系统通知/群聊"""
    if not name:
        return True
    if name in IGNORED_SESSION_NAMES:
        return True
    for tag in IGNORED_PREVIEW_TAGS:
        if tag in raw_text:
            return True
    if IGNORE_GROUP_CHATS and is_group_session(name):
        return True
    return False


# ---------------------------------------------------------------------------
# 窗口控制（零闪烁无感控制 + 原始窗口状态保持）
# ---------------------------------------------------------------------------
def get_weixin_hwnd():
    hwnd = win32gui.FindWindow('Qt51514QWindowIcon', '微信')
    if hwnd == 0:
        hwnd = win32gui.FindWindow('Qt51514QWindowIcon', 'Weixin')
    return hwnd


def make_transparent():
    """把微信主窗口设成全透明（仍在原位，只是看不见）"""
    hwnd = get_weixin_hwnd()
    if not hwnd:
        return
    try:
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style | win32con.WS_EX_LAYERED)
        win32gui.SetLayeredWindowAttributes(hwnd, 0, 0, win32con.LWA_ALPHA)
    except Exception:
        pass


def make_opaque():
    """恢复微信主窗口的不透明度"""
    hwnd = get_weixin_hwnd()
    if not hwnd:
        return
    try:
        win32gui.SetLayeredWindowAttributes(hwnd, 0, 255, win32con.LWA_ALPHA)
    except Exception:
        pass


def stealth_restore():
    """无感恢复窗口（全透明 + 不抢焦点）"""
    hwnd = get_weixin_hwnd()
    if not hwnd:
        return
    make_transparent()

    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
        time.sleep(0.25)
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.25)


def stealth_minimize():
    """最小化窗口，并在最小化后恢复正常不透明度"""
    hwnd = get_weixin_hwnd()
    if not hwnd:
        return
    if not win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        time.sleep(0.25)
    make_opaque()


@contextlib.contextmanager
def stealth_window_context():
    """让 pyweixin 的库调用（dump_sessions / open_dialog_window）全程不可见。

    库里的 Navigator.open_weixin 会 restore 窗口、把尺寸改成 GlobalConfig.window_size
    （默认 1000x1000）、移到屏幕正中、设成置顶、还会抢焦点。

    关键点：**不管窗口原本是否最小化，进入前一律先置成全透明**。只在「已最小化」时
    才透明是不够的 —— 用户把微信开着的时候，库里改尺寸、搬位置的动作会直接被他看见。
    退出时恢复不透明度；若进入时是最小化的则重新最小化。
    （几何与置顶状态的还原由 preserved_window_geometry 负责。）
    """
    hwnd = get_weixin_hwnd()
    was_minimized = bool(hwnd and win32gui.IsIconic(hwnd))

    make_transparent()
    if was_minimized:
        stealth_restore()

    try:
        yield
    finally:
        if was_minimized:
            stealth_minimize()   # 内部会顺手把不透明度恢复回去
        else:
            make_opaque()


@contextlib.contextmanager
def preserved_window_geometry():
    """保住微信主窗口的尺寸/位置/置顶状态。

    pyweixin 的 Navigator.open_weixin 会把窗口 restore 后 SetWindowPos 成
    GlobalConfig.window_size（默认 1000x1000）并移到屏幕正中，同时设成 HWND_TOPMOST，
    而且全库没有任何地方复原。stealth_restore 的全透明只能遮住过程，尺寸/位置/置顶
    是会留下来的副作用 —— 尤其是置顶，会让微信一直浮在所有窗口之上。

    进出本上下文各记录/还原一次即可。
    注意：必须在窗口已恢复（非最小化）时进入，否则读到的几何不可信。
    """
    hwnd = get_weixin_hwnd()
    if not hwnd:
        yield
        return

    try:
        # GetWindowPlacement 返回 (flags, showCmd, ptMin, ptMax, rcNormalPosition)
        # 用 rcNormalPosition 而非 GetWindowRect：最小化/最大化下的 GetWindowRect 不可信
        left, top, right, bottom = win32gui.GetWindowPlacement(hwnd)[4]
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        was_topmost = bool(ex_style & win32con.WS_EX_TOPMOST)
    except Exception:
        left = top = right = bottom = None
        was_topmost = False

    try:
        yield
    finally:
        if left is not None:
            try:
                # 注意：清 WS_EX_TOPMOST 只能靠 SetWindowPos(HWND_NOTOPMOST)，
                # SetWindowLong(GWL_EXSTYLE) 改这一位是无效的（实测验证过）
                cur_topmost = bool(
                    win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOPMOST
                )
                if cur_topmost == was_topmost:
                    # 置顶位无需变动：用 SWP_NOZORDER 还原几何，不扰动 Z 序
                    win32gui.SetWindowPos(
                        hwnd, 0, left, top, right - left, bottom - top,
                        win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER,
                    )
                else:
                    # 置顶位被 open_weixin 改过，只能用 HWND_*TOPMOST 拨回来
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_TOPMOST if was_topmost else win32con.HWND_NOTOPMOST,
                        left, top, right - left, bottom - top,
                        win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER,
                    )
            except Exception:
                pass


@contextlib.contextmanager
def preserved_foreground():
    """拦住 pywinauto 的 set_focus() 把前台窗口从终端抢到微信。

    UIA 的 SetFocus 会连带激活窗口。这里只做「事后补救」，且仅在前台确实被
    微信抢走时才动手 —— 免得跟用户主动切换窗口打架。
    后台进程调 SetForegroundWindow 常被系统拒绝，所以再补一次
    AttachThreadInput 的常规绕过手段。
    """
    wechat = get_weixin_hwnd()
    try:
        before = win32gui.GetForegroundWindow()
    except Exception:
        before = 0

    try:
        yield
    finally:
        # 注意：finally 里不能出现 return，否则会把 with 体内的异常一起吞掉
        # 条件里的 before != wechat 是必要的：微信本来就是前台时不算「被抢走」，
        # 不该去碰前台，也就省掉了无意义的 SetForegroundWindow / AttachThreadInput
        try:
            if (before and wechat and before != wechat
                    and win32gui.GetForegroundWindow() == wechat):
                win32gui.SetForegroundWindow(before)
                if win32gui.GetForegroundWindow() != before:
                    # 常规调用被拒：临时把自己的输入队列挂到当前前台线程上再试一次
                    fg_thread = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[0]
                    cur_thread = win32api.GetCurrentThreadId()
                    if fg_thread and fg_thread != cur_thread:
                        win32process.AttachThreadInput(fg_thread, cur_thread, True)
                        try:
                            win32gui.SetForegroundWindow(before)
                        finally:
                            win32process.AttachThreadInput(fg_thread, cur_thread, False)
        except Exception:
            pass


def get_main_window():
    """获取微信主窗口对象，不改变窗口状态"""
    hwnd = get_weixin_hwnd()
    if hwnd == 0:
        raise RuntimeError('微信未运行，请先启动并登录微信')
    mw = Desktop(backend='uia').window(handle=hwnd)
    cls = mw.class_name()
    if cls == 'mmui::LoginWindow':
        raise RuntimeError('微信未登录，请先登录')
    if cls != 'mmui::MainWindow':
        raise RuntimeError('无法识别微信主界面（需要 UI 树可见的账号）')
    return mw


# ---------------------------------------------------------------------------
# 接收与消息分类过滤
# ---------------------------------------------------------------------------
def classify(cls, txt):
    """把 (class_name, text) 转成 (类型名, 显示内容)"""
    if cls == 'mmui::ChatTextItemView':
        return '文本', txt

    prefix_map = [
        (Special_Labels.File, '文件'),
        (Special_Labels.Image, '图片'),
        (Special_Labels.Video, '视频'),
        (Special_Labels.Link, '链接'),
        (Special_Labels.MiniProgram, '小程序'),
        (Special_Labels.Channels, '视频号'),
        (Special_Labels.RedPacket, '红包'),
        (Special_Labels.Transfer, '转账'),
        (Special_Labels.Emoji, '动画表情'),
    ]
    for label, name in prefix_map:
        tag = label.strip()
        if tag and tag in txt:
            detail = txt.replace(tag, '', 1).strip().split('\n')[0].strip()
            return name, detail
    return '其他', txt


def read_messages(main_window):
    """纯 UIA 读取当前聊天消息列表（自动过滤视频号等屏蔽类型），按时间升序返回。"""
    chat_list = main_window.child_window(**Lists.FriendChatList)
    msgs = []
    for item in chat_list.children():
        cls = item.class_name()
        txt = item.window_text()
        if cls == 'mmui::ChatItemView':  # 时间分隔符，跳过
            continue
        mtype, _ = classify(cls, txt)
        if mtype in IGNORED_MSG_TYPES:   # 过滤视频号卡片消息
            continue
        msgs.append((cls, txt))
    return msgs


def diff_messages(old_msgs, curr_msgs):
    """通过列表尾部切片对齐，提取真正新增的消息"""
    if not old_msgs or old_msgs == curr_msgs:
        return []
    max_k = min(len(old_msgs), len(curr_msgs))
    for k in range(max_k, 0, -1):
        if old_msgs[-k:] == curr_msgs[:k]:
            return curr_msgs[k:]
    return curr_msgs


def list_sessions(main_window, limit=30, filter_official=True):
    """纯 UIA 读取左侧会话列表（可见部分），默认过滤公众号与视频号。"""
    session_list = main_window.child_window(**Main_window.SessionList)
    sessions = []
    for item in session_list.children():
        name = item.automation_id().replace('session_item_', '')
        if not name:
            continue
        raw_txt = item.window_text()
        if filter_official and is_ignored_session(name, raw_txt):
            continue
        m = Regex_Patterns.newMessage_pattern.search(raw_txt)
        new_count = int(m.group(1)) if m else 0
        sessions.append((name, new_count))
    if limit is None:
        return sessions
    return sessions[:limit]


def scan_new_sessions(main_window, current):
    """返回「有未读新消息、≠当前好友、且非公众号/视频号」的 {好友名: 条数}。"""
    result = {}
    for name, count in list_sessions(main_window, limit=None, filter_official=True):
        if name == current:
            continue
        if count > 0:
            result[name] = count
    return result


# ---------------------------------------------------------------------------
# 终端显示
# ---------------------------------------------------------------------------
def clear_screen():
    """清空终端画面。只清命令行的显示内容，不碰任何聊天记录。"""
    try:
        os.system('cls')          # Windows 自带，最稳
    except Exception:
        print('\033[2J\033[H', end='')


# ---------------------------------------------------------------------------
# 主对话循环
# ---------------------------------------------------------------------------
class ChatSession:
    def __init__(self):
        self.current = None               # 当前好友名
        self.history_map = {}             # {好友名: 上次读取的消息列表 [(cls, txt), ...]}
        self.notified = set()             # 已提示过新消息的好友（去重用）
        self.lock = threading.Lock()
        self.stop = threading.Event()
        self.send_x = None                # 发送按钮缓存坐标（客户区相对坐标）
        self.send_y = None
        self.slot_coords = []             # 严格对齐物理行号的坐标列表 [(cx, cy) 或 None, ...]
        self.session_pos = {}             # 按名字缓存的坐标字典 {好友名: (cx, cy)}
        self.last_list = []               # 最近一次列出的会话名列表，用于 /序号 选择
        self.reconnect_failures = 0       # 后台自动重连连续失败次数（防恢复/最小化风暴）

    def _init_coords_stealth(self, main_window):
        """启动时静默采集一次发送按钮坐标和左侧会话坐标（保持原窗口最小化/展开状态）"""
        with stealth_window_context():
            self._refresh_slot_coords(main_window)
            try:
                self._cache_send_btn_pos(main_window)
            except Exception:
                pass

    def _cache_send_btn_pos(self, main_window):
        send_btn = main_window.child_window(**Buttons.SendButton)
        if not send_btn.exists(timeout=1.5):
            raise RuntimeError('未找到发送按钮')
        r = send_btn.rectangle()
        hwnd = get_weixin_hwnd()
        cx, cy = win32gui.ScreenToClient(hwnd, ((r.left + r.right) // 2, (r.top + r.bottom) // 2))
        if cx > 0 and cy > 0:
            self.send_x, self.send_y = cx, cy

    def _refresh_slot_coords(self, main_window):
        """缓存左侧会话列表坐标：
        1. slot_coords 使用 None 占位保持与 list_sessions(filter_official=False) 严格 1:1 行号对齐；
        2. session_pos 按会话名记录有效坐标。
        """
        hwnd = get_weixin_hwnd()
        session_list = main_window.child_window(**Main_window.SessionList)
        coords = []
        pos_map = {}
        for item in session_list.children():
            name = item.automation_id().replace('session_item_', '')
            if not name:
                continue
            r = item.rectangle()
            cx, cy = win32gui.ScreenToClient(hwnd, ((r.left + r.right) // 2, (r.top + r.bottom) // 2))
            if cx > 0 and cy > 0:
                coords.append((cx, cy))
                pos_map[name] = (cx, cy)
            else:
                coords.append(None)  # 严格保留行号槽位占位，防止后续行下标整体错位
        if any(c is not None for c in coords):
            self.slot_coords = coords
            self.session_pos = pos_map

    def _current_chat_name(self, main_window):
        """读取当前主界面聊天对象名（去掉库残留的 title 字段）"""
        criteria = dict(Texts.CurrentChatNameText)
        criteria.pop('title', None)
        t = main_window.child_window(**criteria)
        return t.window_text() if t.exists(timeout=0.5) else None

    def _is_on_friend(self, main_window):
        """校验当前主界面聊天对象是否还是目标好友"""
        try:
            return self._current_chat_name(main_window) == self.current
        except Exception:
            return False

    def _post_click(self, cx, cy):
        """向微信主窗口客户区坐标发送后台点击"""
        hwnd = get_weixin_hwnd()
        lparam = (cy << 16) | (cx & 0xFFFF)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.06)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.35)

    def _click_session_bg(self, main_window, friend):
        """纯后台切换：先按严格对齐的实时行号点击，若失败或为空则按名字缓存坐标尝试"""
        raw_sessions = list_sessions(main_window, limit=None, filter_official=False)
        target_idx = None
        for idx, (name, _) in enumerate(raw_sessions):
            if name == friend:
                target_idx = idx
                break

        # 1) 优先使用严格对齐的物理行号坐标（支持新消息把好友顶到前排后的准确点击）
        if target_idx is not None and target_idx < len(self.slot_coords):
            coord = self.slot_coords[target_idx]
            if coord is not None:
                self._post_click(*coord)
                if self._current_chat_name(main_window) == friend:
                    return True

        # 2) 回退：尝试按名字缓存的坐标（列表未发生重排时有效）
        if friend in self.session_pos:
            self._post_click(*self.session_pos[friend])
            if self._current_chat_name(main_window) == friend:
                return True

        return False

    def _click_session_fg(self, main_window, friend):
        """在恢复状态下点击可见会话项"""
        session_list = main_window.child_window(**Main_window.SessionList)
        for item in session_list.children():
            if item.automation_id() == 'session_item_' + friend:
                item.click_input()
                time.sleep(0.4)
                return self._current_chat_name(main_window) == friend
        return False

    def _print_history(self, friend, msgs, count=DEFAULT_HISTORY_COUNT):
        """打印最近 count 条历史消息"""
        recent = msgs[-count:] if count > 0 else []
        if not recent:
            print(f'  （暂无「{friend}」的可见历史消息）')
            return
        print(f'--- 最近 {len(recent)} 条聊天记录 ---')
        for cls, txt in recent:
            mtype, detail = classify(cls, txt)
            if mtype == '文本':
                print(f'  · {detail}')
            else:
                print(f'  · [{mtype}] {detail}')
        print('------------------------------')

    def _open_friend(self, main_window, friend, show_history=True):
        """切换到指定好友：优先纯后台切换；失败或缺发送坐标时走 stealth_window_context"""
        opened = False
        try:
            if self._click_session_bg(main_window, friend):
                opened = True
        except Exception:
            opened = False

        if not opened or self.send_x is None:
            with stealth_window_context():
                if not opened:
                    for _ in range(2):
                        try:
                            if self._click_session_fg(main_window, friend):
                                opened = True
                                break
                        except Exception:
                            pass
                        time.sleep(0.2)
                if not opened:
                    try:
                        with preserved_window_geometry():
                            Navigator.open_dialog_window(friend=friend, search_pages=15)
                        # 严格校验打开的窗口标题是否等于目标好友，防止误开同名公众号
                        opened = (self._current_chat_name(main_window) == friend)
                    except Exception:
                        opened = False
                if opened:
                    # 几何复原后重新布局，再采集坐标
                    time.sleep(0.3)
                    self._refresh_slot_coords(main_window)
                    try:
                        self._cache_send_btn_pos(main_window)
                    except Exception:
                        pass

        if not opened:
            raise RuntimeError(f'无法打开「{friend}」的会话')

        msgs = read_messages(main_window)
        self.history_map[friend] = msgs
        self.current = friend
        self.reconnect_failures = 0
        self.notified.discard(friend)

        self._note_if_group(main_window, friend)

        if show_history:
            self._print_history(friend, msgs, count=DEFAULT_HISTORY_COUNT)

    def _note_if_group(self, main_window, friend):
        """微信自己的 UIA 能确定当前聊天是不是群聊，认出就永久屏蔽。

        会话列表里群聊和单聊的结构完全一样，按名字猜必然漏，所以这条路才是可靠的：
        刚切进去时顺手查一次顶部有没有群成员数标签（Tools.is_group_chat，
        底层是 Texts.GroupLabelText，一次 0.1s 超时的 UIA 查询）。
        """
        if not IGNORE_GROUP_CHATS or friend in IGNORED_GROUP_NAMES:
            return
        try:
            if Tools.is_group_chat(main_window):
                IGNORED_GROUP_NAMES.add(friend)
                save_blocked_groups()
                print(f'  [群聊] 已屏蔽「{friend}」，之后的会话列表和新消息提醒不再显示它')
        except Exception:
            pass

    def _send_text(self, main_window, text):
        """后台发送：set_focus + set_text 写入 + PostMessage 点击发送按钮"""
        edit_area = main_window.child_window(**Edits.CurrentChatEdit)
        if not edit_area.exists(timeout=2):
            raise RuntimeError('未找到输入框')

        if self.send_x is None:
            with stealth_window_context():
                self._cache_send_btn_pos(main_window)

        if self.send_x is None:
            raise RuntimeError('发送按钮坐标未就绪')

        with preserved_foreground():
            try:
                edit_area.set_focus()
            except Exception:
                pass
        edit_area.set_text(text)
        time.sleep(0.2)

        self._post_click(self.send_x, self.send_y)

    def _match_friend(self, name, main_window):
        """按名字匹配好友：先精确，再前缀；返回匹配到的完整名字或 None"""
        sessions = list_sessions(main_window, limit=None, filter_official=True)
        names = [n for n, _ in sessions]
        if name in names:
            return name
        prefix = [n for n in names if n.startswith(name)]
        if len(prefix) == 1:
            return prefix[0]
        return None

    def _resolve_target(self, text, main_window):
        """把用户输入解析成好友名：支持 /序号 和 /名字"""
        if text.isdigit():
            idx = int(text)
            if 1 <= idx <= len(self.last_list):
                return self.last_list[idx - 1], None
            return None, f'序号 {text} 超出范围（当前列表 {len(self.last_list)} 条），先 /list 查看'
        return (self._match_friend(text, main_window) or text), None

    def _show_list(self, main_window, deep=False, limit=15):
        """列最近会话（自动过滤公众号/视频号/群聊）

        deep=False：纯 UIA 读屏幕上可见的那几行（约 9 条），零打扰、稳定。
        deep=True ：走 Messages.dump_sessions 滚动读取，能凑满 limit 条，但实测不稳定
                    —— 会话列表是虚拟化的 Qt 列表，它内部靠 type_keys 翻页，
                    偶发（连续两次调用出现 41/41/0/0）会整体返回空。所以只给
                    /list all 这种显式请求用，并且失败时退回可见读取。
        """
        if deep:
            failure = None
            with stealth_window_context():
                try:
                    # dump_sessions 内部结尾已自带滚回顶部逻辑，无需额外注入 {HOME}
                    with preserved_window_geometry():
                        raw = Messages.dump_sessions(chat_only=True, close_weixin=False)
                    # 等窗口几何复原后重新布局，再采集坐标（坐标必须对应复原后的尺寸）
                    time.sleep(0.4)
                    self._refresh_slot_coords(main_window)
                    try:
                        self._cache_send_btn_pos(main_window)
                    except Exception:
                        pass
                except Exception as e:
                    # 不要把异常吞成静默的空列表：不然出了问题完全无从下手
                    failure = e
                    raw = []
            entries = []
            for n, t, m in raw:
                preview = str(m).replace('\n', ' ')
                if is_ignored_session(n, preview):
                    continue
                entries.append((n, f'({t})', preview[:25]))
                if len(entries) >= limit:
                    break
        else:
            entries = [(n, f'({c}条新)' if c else '', '')
                       for n, c in list_sessions(main_window, limit=limit, filter_official=True)]

        if not entries and deep:
            # 滚动读取没拿到东西（窗口状态异常等）时退回纯可见读取，别让用户没列表可看
            reason = f'（{type(failure).__name__}: {failure}）' if failure else '（返回空列表）'
            print(f'[提示] 滚动读取没拿到会话{reason}，退回可见会话列表')
            self._show_list(main_window, deep=False, limit=limit)
            return
        if not entries:
            print('未读取到有效好友会话列表')
            return
        self.last_list = [name for name, _, _ in entries]
        print(f'\n最近 {len(entries)} 个会话（已过滤公众号/视频号，输入 /名字 或 /序号 切换）：')
        for i, (name, a, b) in enumerate(entries, 1):
            mark = ' *' if name == self.current else ''
            extra = '  '.join(x for x in (a, b) if x)
            print(f'  [{i:>2}] {name}  {extra}{mark}')

    def run(self):
        GlobalConfig.close_weixin = False
        load_blocked_groups()
        try:
            main_window = get_main_window()
        except RuntimeError as e:
            print(f'[错误] {e}')
            return

        self._init_coords_stealth(main_window)
        self._show_list(main_window, deep=False, limit=15)

        while self.current is None:
            raw = input('\n输入 /名字 或 /序号 选择对话（/list 查看、/clear 清屏、/quit 退出）：').strip()
            if raw.lower() in ('/quit', '/exit', '/q'):
                print('已退出。')
                return
            if raw in ('/clear', '/cls'):
                clear_screen()
                continue
            if raw in ('/list', ''):
                self._show_list(main_window, deep=False, limit=15)
                continue
            if raw == '/list all':
                self._show_list(main_window, deep=True, limit=30)
                continue
            if raw.startswith('/'):
                friend = raw[1:].strip()
                if not friend:
                    continue
                target, err = self._resolve_target(friend, main_window)
                if target is None:
                    print(err)
                    continue
                try:
                    self._open_friend(main_window, target, show_history=True)
                except Exception:
                    print(f'未找到「{friend}」，可用 /list 查看')
                continue

        receiver = threading.Thread(target=self._receiver_loop, args=(main_window,), daemon=True)
        receiver.start()
        print(f'\n已进入与「{self.current}」的对话，输入消息回车发送')
        print('命令：/list 列会话 · /h [条数] 回看历史 · /clear 清屏 · /名字 或 /序号 切换 · /quit 退出\n')

        try:
            while not self.stop.is_set():
                try:
                    text = input('我 > ')
                except (EOFError, KeyboardInterrupt):
                    break
                text = text.strip()
                if text in ('/quit', '/exit', '/q'):
                    break
                if text in ('/clear', '/cls'):
                    with self.lock:
                        clear_screen()
                        print(f'当前对话：{self.current}')
                    continue
                if text == '/list':
                    with self.lock:
                        self._show_list(main_window, deep=False, limit=15)
                    continue
                if text == '/list all':
                    with self.lock:
                        self._show_list(main_window, deep=True, limit=30)
                    continue
                if text == '/h' or text.startswith('/h ') or text == '/history' or text.startswith('/history '):
                    parts = text.split()
                    cnt = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else DEFAULT_HISTORY_COUNT
                    with self.lock:
                        if not self._is_on_friend(main_window):
                            self._open_friend(main_window, self.current, show_history=False)
                        msgs = read_messages(main_window)
                        self.history_map[self.current] = msgs
                        self._print_history(self.current, msgs, count=cnt)
                    continue
                if text.startswith('/'):
                    friend = text[1:].strip()
                    if not friend:
                        continue
                    with self.lock:
                        target, err = self._resolve_target(friend, main_window)
                        if target is None:
                            print(err)
                            continue
                        try:
                            print(f'\n已切换到「{target}」')
                            self._open_friend(main_window, target, show_history=True)
                        except Exception:
                            print(f'未找到「{friend}」，可用 /list 查看')
                    continue
                if not text:
                    continue
                with self.lock:
                    try:
                        if not self._is_on_friend(main_window):
                            self._open_friend(main_window, self.current, show_history=False)
                        self._send_text(main_window, text)
                        self.history_map[self.current] = read_messages(main_window)
                    except Exception as e:
                        print(f'[发送失败] {e}')
        finally:
            self.stop.set()
            receiver.join(timeout=3)
            print('\n已退出对话。')

    def _receiver_loop(self, main_window):
        while not self.stop.is_set():
            time.sleep(1.5)
            try:
                with self.lock:
                    on_friend = self._is_on_friend(main_window)
                    if not on_friend and self.reconnect_failures < MAX_RECONNECT_FAILURES:
                        try:
                            self._open_friend(main_window, self.current, show_history=False)
                            on_friend = True
                        except Exception:
                            self.reconnect_failures += 1
                            if self.reconnect_failures >= MAX_RECONNECT_FAILURES:
                                print(
                                    f'\n[警告] 连续 {MAX_RECONNECT_FAILURES} 次无法切回「{self.current}」，'
                                    f'已暂停自动重连（发消息或 /list 重新选择即可恢复）\n我 > ',
                                    end='', flush=True
                                )

                    if on_friend:
                        self.reconnect_failures = 0
                        curr_msgs = read_messages(main_window)
                        old_msgs = self.history_map.get(self.current, [])
                        new_msgs = diff_messages(old_msgs, curr_msgs)
                        self.history_map[self.current] = curr_msgs

                        for cls, txt in new_msgs:
                            mtype, detail = classify(cls, txt)
                            ts = time.strftime('%H:%M')
                            if mtype == '文本':
                                print(f'\n[{ts}] {self.current}：{detail}\n我 > ', end='', flush=True)
                            else:
                                print(f'\n[{ts}] {self.current}：[{mtype}] {detail}\n我 > ', end='', flush=True)

                    # 熔断/没切回去时也要扫（读的是左侧会话列表，与当前打开的是谁无关），
                    # 否则其他好友的未读提醒会被一并掐断
                    for friend, count in scan_new_sessions(main_window, self.current).items():
                        if friend in self.notified:
                            continue
                        self.notified.add(friend)
                        ts = time.strftime('%H:%M')
                        print(f'\n[{ts}] 【{friend} 发来 {count} 条新消息，输入 /{friend} 切换】\n我 > ', end='', flush=True)
            except Exception:
                pass


def main():
    ChatSession().run()


if __name__ == '__main__':
    main()