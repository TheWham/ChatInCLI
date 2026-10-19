# -*- coding: utf-8 -*-
"""微信命令行摸鱼聊天工具（适配微信 4.x）

在终端里跟多个好友收发消息，全程不操作微信窗口（微信最小化到后台）。
支持多会话快速切换：跟 T 聊天时能识别「肖维发来新消息」，用 /肖维、/T
随时切换对话。

关键实现（微信 4.x 实测）：
- 发送：set_text() 直接写入输入框 + PostMessage 给主窗口发送鼠标点击
  消息（按缓存坐标点击「发送(S)」按钮）。两者都不需要前台/真实鼠标，
  窗口最小化状态下即可完成，彻底不弹 UI。
- 接收：直接读聊天消息列表的 UIA 树（最小化后依然可读）。
- 识别谁发来新消息：读左侧会话列表的「[N条]」未读红标（纯 UIA，后台可读），
  这与「聊天消息层面读不到发件人」是两码事——会话列表层面是有未读标记的。
- 防串聊：收发前用 CurrentChatNameText 校验当前聊天对象，被切走则重开。
"""
import threading
import time
import win32gui
import win32con
from pywinauto import Desktop

from pyweixin import Messages, GlobalConfig
from pyweixin.Uielements import (Lists, Edits, Buttons, Texts, Special_Labels,
                                 Main_window, Regex_Patterns)


# ---------------------------------------------------------------------------
# 窗口控制
# ---------------------------------------------------------------------------
def get_weixin_hwnd():
    hwnd = win32gui.FindWindow('Qt51514QWindowIcon', '微信')
    if hwnd == 0:
        hwnd = win32gui.FindWindow('Qt51514QWindowIcon', 'Weixin')
    return hwnd


def bring_to_front():
    hwnd = get_weixin_hwnd()
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.4)


def minimize():
    hwnd = get_weixin_hwnd()
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
    time.sleep(0.3)


def get_main_window():
    """获取微信主窗口对象，但**不恢复/不改变窗口状态**（避免弹 UI）。
    注意：不用 Navigator.open_weixin，因为它内部会 restore() 把窗口拉到前台。"""
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
# 接收（读取消息）
# ---------------------------------------------------------------------------
def read_messages(main_window):
    """纯 UIA 读取当前聊天消息列表，返回 [(class_name, text), ...]，按时间升序。"""
    chat_list = main_window.child_window(**Lists.FriendChatList)
    msgs = []
    for item in chat_list.children():
        cls = item.class_name()
        txt = item.window_text()
        if cls == 'mmui::ChatItemView':  # 时间分隔符，跳过
            continue
        msgs.append((cls, txt))
    return msgs


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


def list_sessions(main_window, limit=30):
    """纯 UIA 读取左侧会话列表（可见部分），返回 [(名字, 新消息数), ...]。"""
    session_list = main_window.child_window(**Main_window.SessionList)
    sessions = []
    for item in session_list.children():
        name = item.automation_id().replace('session_item_', '')
        if not name:
            continue
        m = Regex_Patterns.newMessage_pattern.search(item.window_text())
        new_count = int(m.group(1)) if m else 0
        sessions.append((name, new_count))
    if limit is None:
        return sessions
    return sessions[:limit]


def scan_new_sessions(main_window, current):
    """返回「有未读新消息且≠当前好友」的 {好友名: 条数}。"""
    result = {}
    for name, count in list_sessions(main_window, limit=None):
        if name == current:
            continue
        if count > 0:
            result[name] = count
    return result


# ---------------------------------------------------------------------------
# 主对话循环
# ---------------------------------------------------------------------------
class ChatSession:
    def __init__(self):
        self.current = None        # 当前好友名
        self.seen_map = {}         # {好友名: 已显示消息签名集合}
        self.notified = set()      # 已提示过新消息的好友（去重用）
        self.lock = threading.Lock()
        self.stop = threading.Event()
        self.send_x = None         # 发送按钮缓存坐标（客户区相对坐标）
        self.send_y = None
        self.session_pos = {}      # 会话列表项坐标缓存 {名字: (cx, cy)}，恢复状态时刷新

    def _get_seen(self, friend):
        if friend not in self.seen_map:
            self.seen_map[friend] = set()
        return self.seen_map[friend]

    def _cache_send_btn_pos(self, main_window):
        send_btn = main_window.child_window(**Buttons.SendButton)
        if not send_btn.exists(timeout=2):
            raise RuntimeError('未找到发送按钮')
        r = send_btn.rectangle()
        hwnd = get_weixin_hwnd()
        cx, cy = win32gui.ScreenToClient(hwnd, ((r.left + r.right) // 2, (r.top + r.bottom) // 2))
        self.send_x, self.send_y = cx, cy

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

    def _refresh_session_pos(self, main_window):
        """刷新会话列表项坐标缓存（需窗口恢复状态下调用，最小化读到的坐标是乱码）"""
        hwnd = get_weixin_hwnd()
        session_list = main_window.child_window(**Main_window.SessionList)
        self.session_pos = {}
        for item in session_list.children():
            name = item.automation_id().replace('session_item_', '')
            if not name:
                continue
            r = item.rectangle()
            cx, cy = win32gui.ScreenToClient(hwnd, ((r.left + r.right) // 2, (r.top + r.bottom) // 2))
            self.session_pos[name] = (cx, cy)

    def _click_session_bg(self, friend):
        """用缓存坐标 PostMessage 点击会话项（无感，最小化下可用）"""
        if friend not in self.session_pos:
            return False
        cx, cy = self.session_pos[friend]
        hwnd = get_weixin_hwnd()
        lparam = (cy << 16) | (cx & 0xFFFF)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.08)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.5)
        return True

    def _open_friend(self, main_window, friend):
        """切换到指定好友。优先后台缓存坐标点击（无感），
        失败则前台刷新坐标缓存后重试（偶发，闪一下）。"""
        opened = False
        # 1) 用缓存坐标后台点击
        for _ in range(3):
            if self._click_session_bg(friend):
                time.sleep(0.6)
                if self._current_chat_name(main_window) == friend:
                    opened = True
                    break
        # 2) 失败（缓存无此人/坐标陈旧），前台刷新缓存后重试
        if not opened:
            bring_to_front()
            time.sleep(0.5)
            self._refresh_session_pos(main_window)
            for _ in range(3):
                if self._click_session_bg(friend):
                    time.sleep(0.6)
                    if self._current_chat_name(main_window) == friend:
                        opened = True
                        break
            minimize()
        if not opened:
            raise RuntimeError(f'无法打开「{friend}」的会话')
        # 3) 基线消息（后台）
        seen = self._get_seen(friend)
        for sig in read_messages(main_window):
            seen.add(sig)
        self.current = friend
        self.notified.discard(friend)

    def _send_text(self, main_window, text):
        """后台发送：set_text 写入 + PostMessage 点击发送按钮，全程最小化"""
        if self.send_x is None:
            # 发送按钮坐标尚未缓存（启动时微信是最小化的），补缓存一次
            bring_to_front()
            try:
                self._cache_send_btn_pos(main_window)
            except Exception:
                pass
            minimize()
        if self.send_x is None:
            raise RuntimeError('未能定位发送按钮')
        edit_area = main_window.child_window(**Edits.CurrentChatEdit)
        if not edit_area.exists(timeout=2):
            raise RuntimeError('未找到输入框')
        edit_area.set_focus()
        edit_area.set_text(text)
        time.sleep(0.25)

        hwnd = get_weixin_hwnd()
        lparam = (self.send_y << 16) | (self.send_x & 0xFFFF)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.08)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.3)

    def _match_friend(self, name, main_window):
        """按名字匹配好友：先精确，再前缀；返回匹配到的完整名字或 None"""
        sessions = list_sessions(main_window)
        names = [n for n, _ in sessions]
        if name in names:
            return name
        prefix = [n for n in names if n.startswith(name)]
        if len(prefix) == 1:
            return prefix[0]
        return None

    def _show_list(self, main_window, deep=False, limit=20):
        """列最近会话。
        deep=False：用最小化下的可见读取（约 9 条，不弹 UI）——启动时用。
        deep=True ：用库方法滚动读取（最多 limit 条，需短暂前台）——/list 用。"""
        if deep:
            bring_to_front()
            try:
                raw = Messages.dump_sessions(chat_only=True, close_weixin=False)[:limit]
            except Exception:
                raw = []
            self._refresh_session_pos(main_window)  # 滚动后坐标变了，刷新缓存
            minimize()
            entries = [(n, f'({t})', str(m).replace('\n', ' ')[:25]) for n, t, m in raw]
        else:
            # 最小化下读可见会话（名字 + 未读条数），不弹 UI
            entries = [(n, f'({c}条新)' if c else '', '')
                       for n, c in list_sessions(main_window, limit=limit)]
        if not entries:
            print('未读取到会话列表')
            return
        print(f'\n最近 {len(entries)} 个会话（/名字 切换）：')
        for i, (name, a, b) in enumerate(entries, 1):
            mark = ' *' if name == self.current else ''
            extra = '  '.join(x for x in (a, b) if x)
            print(f'  [{i:>2}] {name}  {extra}{mark}')

    def run(self):
        GlobalConfig.close_weixin = False
        try:
            main_window = get_main_window()
        except RuntimeError as e:
            print(f'[错误] {e}')
            return
        # 若微信当前可见，直接在恢复状态预缓存坐标（无需任何前台操作，不弹 UI）；
        # 若已最小化，则稍后首次切换/发送时再补缓存
        if not win32gui.IsIconic(get_weixin_hwnd()):
            self._refresh_session_pos(main_window)
            try:
                self._cache_send_btn_pos(main_window)
            except Exception:
                pass
        # 列最近会话（最小化状态也能读，不弹 UI）
        self._show_list(main_window)
        # 初始选择：输入 /名字 或直接回车选择第一个
        while self.current is None:
            raw = input('\n输入 /名字 选择对话（或 /list 查看、/quit 退出）：').strip()
            if raw.lower() in ('/quit', '/exit', '/q'):
                print('已退出。')
                return
            if raw in ('/list', ''):
                self._show_list(main_window)
                continue
            if raw.startswith('/'):
                friend = raw[1:].strip()
                if not friend:
                    continue
                # 前缀匹配拿不到就用原始名字，交给 _open_friend 滚动搜索
                target = self._match_friend(friend, main_window) or friend
                try:
                    self._open_friend(main_window, target)
                except Exception:
                    print(f'未找到「{friend}」，可用 /list 查看')
                continue

        # 若启动时未缓存到发送按钮（当时无聊天打开），进入首个聊天后补缓存（短暂恢复）
        if self.send_x is None:
            bring_to_front()
            try:
                self._cache_send_btn_pos(main_window)
                self._refresh_session_pos(main_window)
            except Exception:
                pass
            minimize()

        receiver = threading.Thread(target=self._receiver_loop, args=(main_window,), daemon=True)
        receiver.start()
        print(f'\n已进入与「{self.current}」的对话，输入消息回车发送')
        print('命令：/list 列会话 · /名字 切换 · /quit 退出\n')

        try:
            while not self.stop.is_set():
                try:
                    text = input('我 > ')
                except (EOFError, KeyboardInterrupt):
                    break
                text = text.strip()
                if text in ('/quit', '/exit', '/q'):
                    break
                if text == '/list':
                    with self.lock:
                        self._show_list(main_window, deep=True)  # 滚动读最近 20 条
                    continue
                if text.startswith('/'):
                    friend = text[1:].strip()
                    if not friend:
                        continue
                    with self.lock:
                        target = self._match_friend(friend, main_window) or friend
                        try:
                            self._open_friend(main_window, target)
                            print(f'已切换到「{target}」')
                        except Exception:
                            print(f'未找到「{friend}」，可用 /list 查看')
                    continue
                if not text:
                    continue
                with self.lock:
                    try:
                        if not self._is_on_friend(main_window):
                            self._open_friend(main_window, self.current)
                        self._send_text(main_window, text)
                        self._get_seen(self.current).add(('mmui::ChatTextItemView', text))
                    except Exception as e:
                        print(f'[发送失败] {e}')
        finally:
            self.stop.set()
            receiver.join(timeout=3)
            # 不恢复窗口，保持最小化（避免退出时弹 UI）
            print('\n已退出对话。')

    def _receiver_loop(self, main_window):
        while not self.stop.is_set():
            time.sleep(2)
            try:
                with self.lock:
                    # 1) 读当前聊天的新消息
                    if not self._is_on_friend(main_window):
                        self._open_friend(main_window, self.current)
                    seen = self._get_seen(self.current)
                    for cls, txt in read_messages(main_window):
                        sig = (cls, txt)
                        if sig in seen:
                            continue
                        seen.add(sig)
                        mtype, detail = classify(cls, txt)
                        ts = time.strftime('%H:%M')
                        if mtype == '文本':
                            print(f'\n[{ts}] {self.current}：{detail}')
                        else:
                            print(f'\n[{ts}] {self.current}：[{mtype}] {detail}')
                    # 2) 扫描会话列表，提示其他好友的新消息
                    for friend, count in scan_new_sessions(main_window, self.current).items():
                        if friend in self.notified:
                            continue
                        self.notified.add(friend)
                        ts = time.strftime('%H:%M')
                        print(f'\n[{ts}] 【{friend} 发来 {count} 条新消息，输入 /{friend} 切换】')
            except Exception:
                # 偶发失败（窗口状态变化），静默跳过，下轮重试
                pass


def main():
    ChatSession().run()


if __name__ == '__main__':
    main()
