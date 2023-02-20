import PySimpleGUI as sg

def get_file(text):
    sg.theme("DarkPurple")
    layout = [[sg.Text(text)],
            [sg.InputText(),sg.FileBrowse()],
            [sg.Submit(), sg.Cancel()]]
    window = sg.Window("Find File",layout)
    event, values=window.read()
    window.close()
    return values[0]




if __name__ == "__main__":
    print(get_file("Get Starfish File"))
    print(get_file("Get Tracker File"))
    