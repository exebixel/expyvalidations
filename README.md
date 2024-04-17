# EXPYVALIDATIONS

[![Pypi](https://img.shields.io/pypi/v/expyvalidations.svg?style=plastic)](https://pypi.org/project/expyvalidations/) [![License](https://img.shields.io/pypi/l/expyvalidations.svg?style=plastic)](https://pypi.org/project/expyvalidations/) [![Python](https://img.shields.io/pypi/pyversions/expyvalidations.svg?style=plastic)](https://pypi.org/project/expyvalidations/) [![Downloads](https://pepy.tech/badge/expyvalidations)](https://pepy.tech/project/expyvalidations) [![Downloads](https://pepy.tech/badge/expyvalidations/month)](https://pepy.tech/project/expyvalidations)

## Description

**ExpyValitations** python library is used to validate data in excel files.
It is possible to validate the type of data, if it is required, if it is unique, perform custom validations and more!

## Installation

```bash
pip install expyvalidations
```

## Usage

```python
# Import the ExpyValidations class
from expyvalidations import ExpyValidations

# Create a new instance of ExpyValidations
ev = ExpyValidations(path_file='path/to/file.xlsx'

# add columns to validate
ev.add_column(key="descricao", name="nome", required=True)
ev.add_column(key="value", name="valor", type="float")

# perform all validations in the file
ev.check_all()

# check if there are errors
if ev.has_errors():
    # get the errors
    print(ev.get_errors())
    """
    [
        {
            "type": "Error",
            "column": "descricao",
            "row": 2,
            "message": "Campo obrigatório"
        },
        {
            "type": "Error",
            "column": "value",
            "row": 3,
            "message": "Value is not a valid number"
        }
    ]
    """
else:
    # get the result
    ev.get_result()
    """
    [
        {
            "descricao": "Produto 1",
            "value": 10.0
        },
        {
            "descricao": "Produto 2",
            "value": 20.0
        }
    ]
    """


```

## License
[Licence MIT](https://github.com/exebixel/expyvalidations/blob/master/LICENSE)

## Documentation
[Click here](https://github.com/exebixel/expyvalidations/wiki) to see the documentation.

## Author

Ezequiel Abreu (Exebixel)

- LinkedIn: [Ezequiel Abreu](https://www.linkedin.com/in/exebixel/)
- GitHub: [Exebixel](https://github.com/exebixel)
- Telegram: [@Exebixel](https://t.me/exebixel)
