# -*- coding: utf-8 -*-
"""可靠的微信消息发送脚本（适配微信 4.x）

问题背景：pyweixin 库自带的 send_messages_to_friend() 使用
pyautogui 的 Ctrl+V 粘贴 + Alt+S 发送，在微信 4.x 富文本输入框里
Ctrl+V 合成按键失效（只打出字母 v），导致消息发不出去。

解决方案：改用 set_text() 直接写入 + 点击「发送(S)」按钮。
"""
import time
import win32gui
from pyweixin import Navigator, GlobalConfig
from pyweixin.Uielements import Edits, Buttons


def _bring_weixin_to_front():
    """把微信窗口置前台，确保键盘/鼠标事件发到微信"""
    hwnd = win32gui.FindWindow('Qt51514QWindowIcon', '微信')
    if hwnd == 0:
        hwnd = win32gui.FindWindow('Qt51514QWindowIcon', 'Weixin')
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.5)


def send_text(friend, text):
    """给指定好友/群聊发送一条文本消息（可靠版）

    Args:
        friend: 好友或群聊的备注名称（需完整精确）
        text: 要发送的文本内容
    """
    if not text or not text.strip():
        raise ValueError('不能发送空白消息')

    GlobalConfig.close_weixin = False
    main_window = Navigator.open_dialog_window(friend=friend)
    _bring_weixin_to_front()

    edit_area = main_window.child_window(**Edits.CurrentChatEdit)
    if not edit_area.exists(timeout=2):
        raise RuntimeError(f'未找到 {friend} 的聊天输入框，请检查备注名是否正确')

    # 直接写入文本（跳过剪贴板+Ctrl+V 这条不可靠路径）
    edit_area.click_input()
    time.sleep(0.2)
    edit_area.set_text(text)
    time.sleep(0.3)

    # 点击发送按钮
    send_btn = main_window.child_window(**Buttons.SendButton)
    if not send_btn.exists(timeout=2):
        raise RuntimeError('未找到发送按钮')
    send_btn.click_input()
    time.sleep(0.8)

    # 校验：发送后输入框应被清空
    if edit_area.window_text().strip():
        raise RuntimeError('发送失败：输入框未清空')
    return True


def send_texts(friend, texts):
    """给指定好友发送多条文本消息"""
    results = []
    for t in texts:
        results.append(send_text(friend, t))
    return results


if __name__ == '__main__':
    # 用法示例：给文件传输助手发消息测试
    send_text('文件传输助手', '可靠的自动发送测试')
    print('发送完成')
