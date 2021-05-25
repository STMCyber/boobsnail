import hashlib
import random
import string
import os
import json
import pathlib
import io

class Boobsnail():
    banner = """
    ___.                ___.     _________             .__.__   
    \_ |__   ____   ____\_ |__  /   _____/ ____ _____  |__|  |  
     | __ \ /  _ \ /  _ \| __ \ \_____  \ /    \\__  \ |  |  |  
     | \_\ (  <_> |  <_> ) \_\ \/        \   |  \/ __ \|  |  |__
     |___  /\____/ \____/|___  /_______  /___|  (____  /__|____/
         \/                  \/        \/     \/     \/         
         Author: @_mzer0 @stm_cyber
         """

    @staticmethod
    def print_banner():
        print(Boobsnail.banner)

def is_number(val):
    if type(val) in (int, float, complex):
        return True
    return False


def md5(val):
    if type(val) == str:
        val = val.encode("utf-8")
    return hashlib.md5(val).digest().hex()

def random_string(length=6):
    letters = string.ascii_lowercase
    return ''.join(random.choice(letters) for i in range(length))

def random_tag(prefix):
    return md5(prefix+random_string())[:16]

def split_blocks(data, block_len):
    for i in range(0, len(data), block_len):
        yield data[i:i+block_len]

def split_data(data, block_len):
    r = []
    for block in split_blocks(data, block_len):
        r.append(block)

    return r

def is_path(p):
    '''
    Checks if path exists
    :param p:
    :return:
    '''
    return os.path.exists(p)

def join_path(p1, p2):
    return os.path.join(p1,p2)

def load_json_file(p):
    '''
    Loads json file and returns dictionary
    :param p:
    :return: dictionary
    '''

    with io.open(p, "r", encoding="utf-8") as json_file:
        data = json_file.read()
    return json.loads(data)

def current_dir():
    return pathlib.Path.cwd()

def write_file(filename, data):
    with io.open(filename, "w", encoding="utf-8") as f:
        f.write(data)


def read_file(filename, mode="r"):
    f = open(filename, mode)
    data = f.read()
    f.close()
    return data