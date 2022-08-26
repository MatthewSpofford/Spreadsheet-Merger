import os
from configparser import ConfigParser
from enum import Enum
from pathlib import Path
from types import DynamicClassAttribute
from typing import Union, Dict


class _ConfigPropValueWrapper:
    def __init__(self, value):
        self.value = value


class ConfigProperty(Enum):
    ORIGINAL_PATH = _ConfigPropValueWrapper("")
    COLUMN_KEY = _ConfigPropValueWrapper("Opportunity Id")
    SHEET_NAME = _ConfigPropValueWrapper("Manager_Opportunity_Dashboard")
    REPLACE_ORIGINAL = _ConfigPropValueWrapper(True)
    MERGED_FILENAME = _ConfigPropValueWrapper("")
    INITIAL_DIR = _ConfigPropValueWrapper("./")

    def __str__(self):
        return self.name

    @DynamicClassAttribute
    def value(self):
        return self._value_.value

    @classmethod
    def generate_default_value_dict(cls):
        default_values = {}
        for prop in cls:
            default_values[prop.name] = prop.value if prop.value is not None else ""
        return default_values


class Config:
    _config_file_location = os.path.join(Path.home(), "spreadsheet-merger.ini")

    _DEFAULT_SECTION = "DEFAULT"
    _DEFAULT_VALUES = ConfigProperty.generate_default_value_dict()

    cp = ConfigParser(defaults=_DEFAULT_VALUES, default_section=_DEFAULT_SECTION, allow_no_value=True)

    _initialized = False

    @classmethod
    def _init(cls):
        if cls._initialized:
           return

        cls.cp.read(cls._config_file_location)
        cls._initialized = True

    @classmethod
    def get(cls, prop: ConfigProperty):
        cls._init()
        if isinstance(prop.value, bool):
            return cls.cp[cls._DEFAULT_SECTION].getboolean(str(prop))
        else:
            return prop.value.__class__(cls.cp[cls._DEFAULT_SECTION][str(prop)])

    @classmethod
    def set(cls, prop: Union[ConfigProperty, Dict[ConfigProperty, str]], value=None):
        cls._init()
        if isinstance(prop, ConfigProperty) and value is not None:
            cls.cp[cls._DEFAULT_SECTION][str(prop)] = str(value)
        elif isinstance(prop, dict):
            for prop_name, prop_value in prop.items():
                cls.cp[cls._DEFAULT_SECTION][str(prop_name)] = str(prop_value)
        else:
            raise ValueError(f"{prop} can only be of type {ConfigProperty}, with a given value as input, or {dict}.")

        cls.save()

    @classmethod
    def save(cls):
        # Write new configuration file
        with open(cls._config_file_location, "w") as config_file:
            cls.cp.write(config_file)
