import wx

from fetch_weather import fetch_weather


class FileDropTarget(wx.FileDropTarget):
    """Drag & Drop Class"""

    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, files):
        # self.window.text_entry.SetLabel(files[0])
        fetch_weather(files[0])
        return 0


class App(wx.Frame):
    """GUI"""

    def __init__(self, parent, id, title):
        wx.Frame.__init__(
            self, parent, id, title, size=(500, 300), style=wx.DEFAULT_FRAME_STYLE
        )

        # パネル
        p = wx.Panel(self, wx.ID_ANY)

        label = wx.StaticText(
            p, wx.ID_ANY, "ここにファイルをドロップしてください", style=wx.SIMPLE_BORDER | wx.TE_CENTER
        )
        label.SetBackgroundColour("#e0ffff")

        # ドロップ対象の設定
        label.SetDropTarget(FileDropTarget(self))

        # テキスト入力ウィジット
        # self.text_entry = wx.TextCtrl(p, wx.ID_ANY)

        button = wx.Button(p, wx.ID_ANY, "ファイル選択")
        button.Bind(wx.EVT_BUTTON, self.open_file)

        # レイアウト
        layout = wx.BoxSizer(wx.VERTICAL)
        layout.Add(label, flag=wx.EXPAND | wx.ALL, border=10, proportion=2)
        layout.Add(button, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)
        # layout.Add(self.text_entry, flag=wx.EXPAND | wx.ALL, border=10)
        p.SetSizer(layout)

        self.Show()

    def open_file(self, event):
        filters = "エクセルファイル(*.xlsx,*.xls)|*.xlsx;*.xls"
        # filetype = """\
        #     TXTfiles(*.txt)|*.txt|
        #     All file(*)|*"""
        with wx.FileDialog(
            None, u"対象のファイル選択してください", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        ) as dialog:
            dialog.SetWildcard(filters)
            # ファイルが選択されたとき
            if dialog.ShowModal() == wx.ID_OK:
                # 選択したファイルパスを取得する
                # self.text_entry.SetLabel(dialog.GetPath())
                fetch_weather(dialog.GetPath())


def main():
    app = wx.App()
    App(None, -1, "天気情報ゲット♪")
    app.MainLoop()


if __name__ == "__main__":
    main()
