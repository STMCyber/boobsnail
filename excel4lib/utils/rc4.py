#!/usr/bin/env python
# -*- coding: utf-8 -*-
# author: @manojpandey
# Copied from https://github.com/manojpandey/rc4
from .utils import *

import codecs

MOD = 256

class RC4(object):
    '''
    Allows to encrypt/decrypt text by using RC4 encryption

    Usage:
    1. Generate key
    2. Encrypt text
    ```
    key = RC4.get_key(10)
    cipher_text = RC4.encrypt("TEST", key)
    ```

    '''
    @staticmethod
    def get_key(length=10):
        return random_string(length)

    @staticmethod
    def KSA(key):
        '''
        Key Scheduling Algorithm (from wikipedia):
        ```
            for i from 0 to 255
                S[i] := i
            endfor
            j := 0
            for i from 0 to 255
                j := (j + S[i] + key[i mod keylength]) mod 256
                swap values of S[i] and S[j]
            endfor
        ```
        '''
        key_length = len(key)
        # create the array "S"
        S = list(range(MOD))  # [0,1,2, ... , 255]
        j = 0
        for i in range(MOD):
            j = (j + S[i] + key[i % key_length]) % MOD
            S[i], S[j] = S[j], S[i]  # swap values

        return S

    @staticmethod
    def PRGA(S):
        '''
        Psudo Random Generation Algorithm (from wikipedia):
        ```
            i := 0
            j := 0
            while GeneratingOutput:
                i := (i + 1) mod 256
                j := (j + S[i]) mod 256
                swap values of S[i] and S[j]
                K := S[(S[i] + S[j]) mod 256]
                output K
            endwhile
        ```
        '''
        i = 0
        j = 0
        while True:
            i = (i + 1) % MOD
            j = (j + S[i]) % MOD

            S[i], S[j] = S[j], S[i]  # swap values
            K = S[(S[i] + S[j]) % MOD]

            yield K

    @staticmethod
    def get_keystream(key):
        ''' Takes the encryption key to get the keystream using PRGA

            return object is a generator
        '''
        # For plaintext key, use this
        key = [ord(c) for c in key]
        S = RC4.KSA(key)
        return RC4.PRGA(S)

    @staticmethod
    def encrypt_logic(key, text):
        ''' :key -> encryption key used for encrypting, as hex string

            :text -> array of unicode values/ byte string to encrpyt/decrypt
        '''

        # If key is in hex:
        # key = codecs.decode(key, 'hex_codec')
        # key = [c for c in key]
        keystream = RC4.get_keystream(key)

        res = []
        for c in text:
            val = ("%04X" % (c ^ next(keystream)))  # XOR and taking hex
            res.append(val)
        return ''.join(res)

    @staticmethod
    def encrypt(key, plaintext):
        ''' :key -> encryption key used for encrypting, as hex string

            :plaintext -> plaintext string to encrpyt
        '''
        plaintext = [ord(c) for c in plaintext]
        return RC4.encrypt_logic(key, plaintext)

    @staticmethod
    def decrypt(key, ciphertext):
        ''' :key -> encryption key used for encrypting, as hex string

            :ciphertext -> hex encoded ciphered text using RC4
        '''
        #ciphertext = codecs.decode(ciphertext, 'hex_codec')
        c = []
        for x in range(0, len(ciphertext), 4):
            c.append(int(ciphertext[x:x+4], 16))
        res = RC4.encrypt_logic(key, c)
        r = ""
        for y in range(0, len(res), 4):
            r = r + chr(int(res[y:y+4], 16))
        #return codecs.decode(res, 'hex_codec')
        return r


    @staticmethod
    def encrypt_logic_ks(keystream, text):
        ''' :keystream -> encryption keystream used for encrypting, as hex string

            :text -> array of unicode values/ byte string to encrpyt/decrypt
        '''
        res = []
        for c in text:
            val = ("%04X" % (c ^ next(keystream)))  # XOR and taking hex
            res.append(val)
        return ''.join(res)

    @staticmethod
    def encrypt_ks(keystream, plaintext):
        ''' :keystream -> encryption keystream used for encrypting, as hex string

            :plaintext -> plaintext string to encrpyt
        '''
        plaintext = [ord(c) for c in plaintext]
        return RC4.encrypt_logic_ks(keystream, plaintext)

    @staticmethod
    def decrypt_ks(keystream, ciphertext):
        ''' :keystream -> encryption keystream used for encrypting, as hex string

            :ciphertext -> hex encoded ciphered text using RC4
        '''
        #ciphertext = codecs.decode(ciphertext, 'hex_codec')
        c = []
        for x in range(0, len(ciphertext), 4):
            c.append(int(ciphertext[x:x+4], 16))
        res = RC4.encrypt_logic_ks(keystream, c)
        #return codecs.decode(res, 'hex_codec')
        r = ""
        for y in range(0, len(res), 4):
            r = r + chr(int(res[y:y+4], 16))
        #return codecs.decode(res, 'hex_codec')
        return r