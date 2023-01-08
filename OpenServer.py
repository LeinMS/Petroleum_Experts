import win32com.client


class OpenServer:
    def __init__(self):
        self.status = 'disconnected'
        self.server = None
        self.error = ''

    def __exit__(self, *args):
        self.disconnect()

    def __enter__(self):
        self.connect()
        return self

    def connect(self):
        self.server = win32com.client.Dispatch("PX32.OpenServer.1")
        self.status = 'connected'
        print('Соединение с OpenServer установлено')

    def disconnect(self):
        self.server = None
        self.status = 'disconnected'
        print('Соединение с OpenServer разорвано')

    def get_app_name(self, os_str):
        return os_str.split('.')[0].upper()

    def get_value(self, os_str):
        try:
            value = self.server.GetValue(os_str)
            app_name = self.get_app_name(os_str)
            err = self.server.GetLastError(app_name)
            if err > 0:
                self.error = self.server.GetLastErrorMessage(app_name)
                raise SyntaxError(self.error)
            return value
        except SyntaxError:
            self.disconnect()
            raise

    def set_value(self, os_str, val):
        try:
            err = self.server.SetValue(os_str, val)
            app_name = self.get_app_name(os_str)
            err = self.server.GetLastError(app_name)
            if err > 0:
                self.error = self.server.GetLastErrorMessage(app_name)
                raise SyntaxError(self.error)
        except SyntaxError:
            self.disconnect()
            raise

    def do_command(self, os_str):
        try:
            err = self.server.docommand(os_str)
            app_name = self.get_app_name(os_str)
            if err > 0:
                self.error = self.server.GetLastErrorMessage(app_name)
                raise SyntaxError(self.error)
        except SyntaxError:
            self.disconnect()
            raise
