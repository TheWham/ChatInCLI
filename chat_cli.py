# -*- coding: utf-8 -*-
"""微信命令行摸鱼聊天工具（适配微信 4.x）

在终端里选取一个最近会话，跟对方收发消息，全程不操作微信窗口
（微信最小化到后台）。仅支持单聊、文本 + 文件/图片提示。

关键实现（微信 4.x 实测）：
- 发送：set_text() 直接写入输入框 + PostMessage 给主窗口发送鼠标点击
  消息（按缓存坐标点击「发送(S)」按钮）。两者都不需要前台/真实鼠标，
  窗口最小化状态下即可完成，彻底不弹 UI。
  （微信 4.x 的 Qt 发送按钮不支持 UIA InvokePattern，invoke() 无效，
   但 PostMessage 的 WM_LBUTTONDOWN/UP 可以触发发送。）
- 接收：直接读聊天消息列表的 UIA 树（最小化后依然可读）。单聊场景下
  「凡是我没主动发过的新消息，就是对方发的」，记录自己发过的内容即可
  判断来源。
- 防串聊：收发前用 CurrentChatNameText 校验当前聊天对象，若被微信因
  新消息自动切走，则重新打开目标会话，避免发错人/读错人。
"""
import threading
import time
import win32gui
import win32con

from pyweixin import Navigator, Messages, GlobalConfig
from pyweixin.Uielements import Lists, Edits, Buttons, Texts, Special_Labels


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


# ---------------------------------------------------------------------------
# 接收（读取消息）
# ---------------------------------------------------------------------------
def read_messages(main_window):
    """纯 UIA 读取聊天消息列表，返回 [(class_name, text), ...]，按时间升序。
    最小化后依然可读，无需键盘/前台。
    """
    chat_list = main_window.child_window(**Lists.FriendChatList)
    msgs = []
    for item in chat_list.children():
        cls = item.class_name()
        txt = item.window_text()
        if cls == 'mmui::ChatItemView':  # 时间分隔符，跳过
            continue
        msgs.append((cls, txt))
    return msgs


def get_last_message_item(main_window):
    """返回最后一条非时间分隔符的消息元素（用于判断发件人）"""
    chat_list = main_window.child_window(**Lists.FriendChatList)
    for item in reversed(chat_list.children()):
        if item.class_name() != 'mmui::ChatItemView':
            return item
    return None


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


# ---------------------------------------------------------------------------
# 会话选择
# ---------------------------------------------------------------------------
def select_friend(n=50):
    print('正在读取最近会话列表...')
    sessions = Messages.dump_sessions(chat_only=True, close_weixin=False)
    sessions = sessions[:n]
    if not sessions:
        raise RuntimeError('未读取到任何会话')
    print(f'\n最近 {len(sessions)} 个会话：')
    for i, (name, t, msg) in enumerate(sessions, 1):
        print(f'  [{i:>2}] {name}  ({t})  {msg[:30]}')
    while True:
        raw = input('\n选择编号（q 退出）：').strip()
        if raw.lower() in ('q', 'quit', 'exit'):
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(sessions):
            return sessions[int(raw) - 1][0]
        print('无效输入，请重新输入编号')


# ---------------------------------------------------------------------------
# 主对话循环
# ---------------------------------------------------------------------------
class ChatSession:
    def __init__(self, friend):
        self.friend = friend
        self.lock = threading.Lock()
        self.stop = threading.Event()
        self.seen = set()          # 已显示过的消息签名 (class, text)
        self.send_x = None         # 发送按钮缓存坐标（客户区相对坐标）
        self.send_y = None

    def _cache_send_btn_pos(self, main_window):
        """缓存发送按钮中心点相对客户区的坐标。必须在窗口非最小化时调用。"""
        send_btn = main_window.child_window(**Buttons.SendButton)
        if not send_btn.exists(timeout=2):
            raise RuntimeError('未找到发送按钮')
        r = send_btn.rectangle()
        hwnd = get_weixin_hwnd()
        cx, cy = win32gui.ScreenToClient(hwnd, ((r.left + r.right) // 2, (r.top + r.bottom) // 2))
        self.send_x, self.send_y = cx, cy

    def _is_on_friend(self, main_window):
        """校验当前主界面聊天对象是否还是目标好友"""
        try:
            name_text = main_window.child_window(**Texts.CurrentChatNameText)
            return name_text.exists(timeout=0.5) and name_text.window_text() == self.friend
        except Exception:
            return False

    def _reopen(self, main_window):
        """聊天被切走时，重新打开目标会话（需短暂前台）"""
        bring_to_front()
        # 用 open_dialog_window（会话列表滚动定位）而非 search_friend：
        # search_friend 走顶部搜索栏，在微信 4.x 下搜索框元素偶发解析为空列表而报错
        Navigator.open_dialog_window(friend=self.friend, search_pages=10)
        self._cache_send_btn_pos(main_window)
        minimize()

    def _send_text(self, main_window, text):
        """后台发送：set_text 写入 + PostMessage 点击发送按钮，全程最小化"""
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

    def run(self):
        GlobalConfig.close_weixin = False
        main_window = Navigator.open_dialog_window(friend=self.friend)

        # 窗口此时处于恢复状态，缓存发送按钮坐标；并把已有消息标记为已见
        self._cache_send_btn_pos(main_window)
        for sig in read_messages(main_window):
            self.seen.add(sig)

        receiver = threading.Thread(target=self._receiver_loop, args=(main_window,), daemon=True)
        receiver.start()

        minimize()
        print(f'\n已进入与「{self.friend}」的对话，输入消息回车发送，/quit 退出\n')

        try:
            while not self.stop.is_set():
                try:
                    text = input('我 > ')
                except (EOFError, KeyboardInterrupt):
                    break
                text = text.strip()
                if text in ('/quit', '/exit', '/q'):
                    break
                if not text:
                    continue
                with self.lock:
                    try:
                        if not self._is_on_friend(main_window):
                            self._reopen(main_window)
                        self._send_text(main_window, text)
                        self.seen.add(('mmui::ChatTextItemView', text))
                    except Exception as e:
                        print(f'[发送失败] {e}')
        finally:
            self.stop.set()
            receiver.join(timeout=3)
            bring_to_front()  # 退出时恢复窗口
            print('\n已退出对话。')

    def _receiver_loop(self, main_window):
        while not self.stop.is_set():
            time.sleep(2)
            try:
                with self.lock:
                    if not self._is_on_friend(main_window):
                        # 聊天被切走，重新打开目标会话，本轮跳过
                        self._reopen(main_window)
                        continue
                    for cls, txt in read_messages(main_window):
                        sig = (cls, txt)
                        if sig in self.seen:
                            continue
                        self.seen.add(sig)
                        mtype, detail = classify(cls, txt)
                        ts = time.strftime('%H:%M')
                        if mtype == '文本':
                            print(f'\n[{ts}] {self.friend}：{detail}')
                        else:
                            print(f'\n[{ts}] {self.friend}：[{mtype}] {detail}')
            except Exception:
                # 偶发失败（窗口状态变化），静默跳过，下轮重试
                pass


def main():
    friend = select_friend(50)
    if friend is None:
        print('已取消。')
        return
    ChatSession(friend).run()


if __name__ == '__main__':
    main()
