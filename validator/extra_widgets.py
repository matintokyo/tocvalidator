import wx
import os


class FileDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        super().__init__()
        self.window: wx.FilePickerCtrl = window

    def OnDropFiles(self, x, y, filenames):
        if filenames:
            file_path = filenames[0]
            if os.path.isfile(file_path):
                self.window.SetPath(file_path)
                self.window.GetTextCtrl().SetInsertionPointEnd()
                wx.PostEvent(self.window.GetEventHandler(),
                             wx.PyCommandEvent(wx.EVT_FILEPICKER_CHANGED.typeId, self.window.GetId()))

        return True


class DropableFilePickerCtrl(wx.FilePickerCtrl):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.drop_target = FileDropTarget(self)
        self.SetDropTarget(self.drop_target)


class DirDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        super().__init__()
        self.window: wx.DirPickerCtrl = window

    def OnDropFiles(self, x, y, filenames):
        if filenames:
            dir_path = filenames[0]
            if os.path.isdir(dir_path):
                self.window.SetPath(dir_path)
                self.window.GetTextCtrl().SetInsertionPointEnd()
                wx.PostEvent(self.window.GetEventHandler(),
                             wx.PyCommandEvent(wx.EVT_DIRPICKER_CHANGED.typeId, self.window.GetId()))

        return True


class DropableDirPickerCtrl(wx.DirPickerCtrl):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.drop_target = DirDropTarget(self)
        self.SetDropTarget(self.drop_target)
