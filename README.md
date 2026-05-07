# OpenGL Text Spinner

An OpenGL picker app for entering names or text values, spinning them, and randomly stopping on one winner.

## Setup

Install Python 3.10 or newer, then install the dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run

```powershell
python spinner_app.py
```

Or double-click `run.bat`.

This version uses `pyglet`, which avoids pygame build issues on newer Python versions.

## Controls

- Type in the input box and press `Enter` or click `Add` to add a name/text value.
- Click `Upload Excel` to select an `.xlsx` file. The file should have only a column; that column becomes the values list.
- Click an item in the list to select it.
- Click `Remove` to delete the selected item.
- Click `Start` to spin.
- Click `Reset` to stop the spin and return the wheel to its initial position.
- Press `Esc` to quit.
