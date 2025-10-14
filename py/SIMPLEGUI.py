import FreeSimpleGUI as sg

# ------------------------------
# Theme & Styles
# ------------------------------
sg.set_options(font=("Helvetica", 11))

sg.LOOK_AND_FEEL_TABLE["GameBoy"] = {
    "BACKGROUND": "#C0C0C0",
    "TEXT": "#0F380F",
    "INPUT": "#C0C0C0",
    "TEXT_INPUT": "#0F380F",
    "SCROLL": "#0F380F",
    "BUTTON": ("#FFFFFF", "#808080"),
    "PROGRESS": ("#0F380F", "#8BAC0F"),
    "BORDER": 1,
    "SLIDER_DEPTH": 0,
    "PROGRESS_DEPTH": 0,
}

sg.theme("GameBoy")

btn_size = (8, 3)
dpad_button_style = {
    "button_color": ("#FFFFFF", "#505050"),
    "border_width": 3,
    "pad": (0, 0),
}
side_button_style = {
    "button_color": ("#FFFFFF", "#C04080"),
    "border_width": 3,
    "pad": (0, 0),
}
bottom_button_style = {
    "button_color": ("#000000", "#808080"),
    "border_width": 3,
    "pad": (0, 0),
}


# ------------------------------
# Layout
# ------------------------------
def create_window():
    dpad_layout = [
        [
            sg.Text(
                "Suetendo GAME BOY™",
                justification="left",
                font=("Helvetica", 14, "bold"),
            )
        ],
        [sg.Text("", size=(1, 1))],
        [sg.Button("Sort Renewal List", size=btn_size, **dpad_button_style)],
        [
            sg.Button("Manual Renewal Letter", size=btn_size, **dpad_button_style),
            sg.Text("", size=(4, 1)),
            sg.Button("ICBC E-Stamp Tool", size=btn_size, **dpad_button_style),
        ],
        [
            sg.Text("", size=(1, 1)),
            sg.Button("Instant CSIO", size=btn_size, **dpad_button_style),
            sg.Text("", size=(1, 1)),
        ],
        [sg.Text("", size=(1, 1))],
    ]

    layout = [
        [
            sg.Multiline(
                default_text="Welcome!\nPress buttons to interact.",
                size=(40, 12),
                key="-SCREEN-",
                background_color="#8BAC0F",
                text_color="#000000",
                disabled=True,
                border_width=2,
                pad=(0, 5),
            )
        ],
        [
            sg.Column(dpad_layout, element_justification="center", pad=(5, 0)),
            sg.Column(
                [
                    [
                        sg.Button(
                            "Copy & Rename ICBC", size=btn_size, **side_button_style
                        ),
                        sg.Button(
                            "Auto Renewal Letter", size=btn_size, **side_button_style
                        ),
                    ],
                    [
                        sg.Text("B", size=btn_size, justification="center"),
                        sg.Text("A", size=btn_size, justification="center"),
                    ],
                ],
                element_justification="center",
                pad=(15, 0),
            ),
        ],
        [
            sg.Button("Reconciler", size=(15, 1), **bottom_button_style),
            sg.Button("Exit", size=(18, 1), **bottom_button_style),
        ],
        [
            sg.Text("Select", size=(6, 1), justification="center"),
            sg.Text("Start", size=(6, 1), justification="center"),
        ],
    ]

    return sg.Window(
        "Game Boy Emulator", layout, element_justification="center", finalize=True
    )


# ------------------------------
# Run GUI with callback functions
# ------------------------------
def run(gui_config):
    """
    gui_config: dict mapping button names to functions
    Example: {'Sort Renewal List': sort_renewal_list, 'ICBC E-Stamp': icbc_estamp}
    """
    window = create_window()
    screen = window["-SCREEN-"]

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Exit"):
            break

        screen.update(screen.get() + f"\nButton pressed: {event}")

        # Call the corresponding function if it exists
        if event in gui_config:
            try:
                result = gui_config[event]()  # Call the function
                if result is not None:
                    screen.update(screen.get() + f"\n{result}")
            except Exception as e:
                screen.update(screen.get() + f"\nError: {e}")

    window.close()
