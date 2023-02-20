import PySimpleGUI as sg


def get_file(text,initial_folder="Downloads"):
    sg.theme("DarkPurple")
    layout = [[sg.Text(text)],
              [sg.InputText(), sg.FileBrowse(initial_folder=initial_folder)],
              [sg.Submit(), sg.Cancel()]]
    window = sg.Window("Find File", layout)
    event, values = window.read()
    window.close()
    return values[0]

def which_subj(text):
    sg.theme("DarkPurple")
    layout = [[sg.Text(text)],
                [sg.Button('MATH'),sg.Button('ENGL')]]
    window = sg.Window("Subject", layout)
    event, values = window.read()
    window.close()
    return event

def yes_no(text):
    sg.theme("DarkPurple")
    layout = [[sg.Text(text)],
                [sg.Button('Yes'),sg.Button('No')]]
    window = sg.Window("Yes or No", layout)
    event, values = window.read()
    window.close()
    return event

if __name__ == "__main__":
    print(yes_no("choose one"))