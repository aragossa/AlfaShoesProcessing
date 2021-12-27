import os
from configparser import ConfigParser


def read_config(section):
    filename = os.path.join("Utils", "ConfigReader", "config.ini")
    parser = ConfigParser()
    parser.read(filename)
    config = {}
    if parser.has_section(section):
        for param in parser.items(section):
            config[param[0]] = param[1]
    else:
        raise Exception("Section {0} not found in the {1} file".format(section, filename))
    return config
