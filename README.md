# MnémoChoice

MnémoChoice is a Python written tool that streamlines word selection by providing quick, standardized shortcuts for commonly used terms in programming. Designed to enhance productivity and standardization, it helps developers instantly retrieve and insert predefined word choices with just a few keystrokes.

## Installation

Use the package manager [pip](https://pip.pypa.io/en/stable/) to install the packages.

```bash
pip install pandas tk psutil shutil keyboard pynput configparser pyperclip openpyxl unicodedata
```

## Usage

Edit the config.ini file according to your needs specifying : 
```bash
[FILE]
PATH = path to your file over network or locally
TAB = sheet name on which is data
COLUMN1 = column name for complete words
COLUMN2 = column name for words shortcut

[KEYBOARD]
SHORTCUT = shortcut to use to launch the utility

[PROCESS]
autokill = 0 to keep window open after selection or 1 to close automatically

[UI]
opacity = 0 to 1 → 0 completely transparent and 1 completely opaque
```

Launch the executable

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)