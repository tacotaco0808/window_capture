import tkinter as tk
import pygetwindow as gw
import pyautogui as pag
import re
import time
import win32con  # win32guiはpywin32だった
import win32gui
import my_icon  # ウィンドウアイコン作成
from PIL import ImageGrab, ImageTk


def get_window_handle(window_title, index):
    """タイトル部分一致ウィンドウのハンドルを取得
    :param window_title: 検索するウィンドウのタイトル
    :param index: 複数ウィンドウの際に番号を指定
    :return: ウィンドウのハンドルを返す、見つからない場合0
    """

    tmp = gw.getWindowsWithTitle(window_title)
    if len(window_title) == 0:  # title入力が空の時エラーを返させる
        tmp = False

    if tmp:
        handle_list = re.findall(r"\d+", str(tmp))
        list_size = len(handle_list)

        # rangeはリストのサイズ÷２
        for num in range(int(list_size / 2)):
            handle_list.remove("32")

        # 指定したindexの変数が存在しない時の例外処理
        try:
            return handle_list[index]
        except IndexError as e:
            return 0
    else:  # リストが空ならエラー
        return 0


def show_window(window_handle):
    """指定したウィンドウを表示
    :param window_handle: ウィンドウのハンドル
    :return: なし
    """
    if win32gui.IsIconic(window_handle):  # 最小化されているか
        # 最小化解除
        win32gui.ShowWindow(window_handle, 1)
        # ウィンドウを最前面に固定
        win32gui.SetWindowPos(window_handle, win32con.HWND_TOPMOST,
                              0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    else:
        win32gui.SetWindowPos(window_handle, win32con.HWND_TOPMOST,
                              0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    # 最前面の固定解除
    win32gui.SetWindowPos(window_handle, win32con.HWND_NOTOPMOST,
                          0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)


class App(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)  # 継承した Frame の __init__() の呼び出し

        # get_window_handleで取得したウィンドウの名前
        self.window_name = None

        # create_toolbarで使う変数
        self.entry_text = tk.StringVar()
        self.spin_box = None
        self.spinbox_val = tk.StringVar()
        self.spinbox_val.set("0")

        # screenshotした画像を表示するための変数
        self.canvas = None
        self.photo = None

        # 本アプリのウィンドウ設定
        self.master.title("RealTimeCapture")
        photo = my_icon.get_icon()  # アイコン作成
        self.master.iconphoto(False, photo)
        scr_w, scr_h = pag.size()  # PCサイズ取得
        scr_w = int(scr_w / 4)
        scr_h = int(scr_h / 4)
        self.master.geometry(str(scr_w) + "x" + str(scr_h))
        self.master.attributes("-topmost", True)

        # ツールバー作成
        self.create_toolbar()

        # 画像表示領域作成・配置
        self.image_frame = tk.Frame(self.master)
        self.image_frame.pack()
        self.image_frame.winfo_width()
        self.image_canvas = tk.Canvas(self.master, background="#008080")
        self.image_canvas.pack(expand=True, fill=tk.BOTH)

    def create_toolbar(self):
        """ウィンドウ上部のツールバー作成
        :return: なし
        """

        # rootウィンドウにtool_barフレームを作成
        tool_bar = tk.Frame(self.master, borderwidth=5, relief=tk.RAISED, width=500)
        tool_bar.pack(fill=tk.X)

        # searchボタン作成・配置
        btn_search = tk.Button(  # buttonのcommand指定の際に引数が必要なときは必ず、lambdaで記述する<-正常に動かない
            tool_bar, text="Search",
            command=lambda: self.upload_window_image(str(self.entry_text.get()), int(self.spinbox_val.get()))
        )
        btn_search.pack(side=tk.LEFT, )

        # スピンボックスの作成・配置
        self.spin_box = tk.Spinbox(
            tool_bar,
            width=5,
            format='%2.0f',
            state='readonly',
            textvariable=self.spinbox_val,
            from_=0,
            to=9,
            increment=1,
        )
        self.spin_box.pack(side=tk.RIGHT, )

        # テキストボックス(Entry)作成・配置
        entry = tk.Entry(tool_bar,
                         width=50,
                         textvariable=self.entry_text

                         )
        entry.pack(side=tk.LEFT, fill=tk.X)

    def upload_window_image(self, window_title, index):
        """Searchボタンが押された際に画像を描画
        :param str window_title:検索するウィンドウのタイトル
        :param int index:複数ウィンドウの際に番号を指定
        :return:なし
        """

        # 描画するキャンバスのサイズを取得
        canvas_width = self.image_canvas.winfo_width()
        canvas_height = self.image_canvas.winfo_height()

        # ウィンドウのハンドルが見つかる場合に描画処理
        handle = get_window_handle(window_title, index)
        if handle != 0:
            handle_screenshot = self.screenshot(handle)
            self.create_canvas_img(handle_screenshot, canvas_width, canvas_height)

        else:
            print("一致するウィンドウがないよ!!")

    # スクリーンショットをとる
    def screenshot(self, window_handle):
        """スクリーンショットをとる
        :param window_handle: ウィンドウのハンドル
        :return: PILでとったスクリーンショットを返す
        """
        # 本アプリウィンドウを非表示
        self.master.withdraw()
        show_window(window_handle)
        # ウィンドウを最大化
        win32gui.ShowWindow(window_handle, win32con.SW_MAXIMIZE)
        time.sleep(1)
        # 全画面スクリーンショット
        image = ImageGrab.grab()
        # ウィンドウを最小化
        win32gui.ShowWindow(window_handle, win32con.SW_MINIMIZE)
        self.master.deiconify()  # 非表示にしていた本アプリのウィンドウを再表示

        return image

    def create_canvas_img(self, image, width, height):
        """指定サイズで画像をリサイズし、キャンバスに描画
        :param image: PILでとったスクリーンショット
        :param width: 幅
        :param height: 高さ
        :return: なし
        """
        # リサイズ
        image = image.resize((int(width), int(height)))
        # PILでとったスクリーンショットをキャンバスへ描画するために変換
        self.photo = ImageTk.PhotoImage(image)
        self.image_canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)


if __name__ == '__main__':
    root = tk.Tk()
    f1 = App(master=root)
    f1.mainloop()
