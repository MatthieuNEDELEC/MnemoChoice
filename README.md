# MnémoChoice

MnémoChoice is a Python tool that streamlines word selection by providing quick, standardized shortcuts for commonly used terms in programming. Designed to enhance productivity, it helps developers instantly retrieve and insert predefined word choices with just a few keystrokes.

## Installation

Use the package manager [pip](https://pip.pypa.io/en/stable/) to install the packages.

```bash
pip install pandas tkinter psutil keyboard configparser
```

## Usage

Edit the config.ini file according to your needs pecifying : 
```bash
[FILE]
PATH = path to your file over network or locally
TAB = sheet name on which is data
COLUMN1 = column name for complete words
COLUMN2 = column name for words shortcut

[KEYBOARD]
SHORTCUT = shortcut to use to launch the utility
```

Laucnh the executable (currently only Windows is supported due to killing process concerns)

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

[MIT](https://choosealicense.com/licenses/mit/)