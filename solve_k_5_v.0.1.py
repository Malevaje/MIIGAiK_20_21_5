# -*- coding: utf-8 -*-
"""
Created on Thu Aug 20 11:35:40 2020

@author: 1
"""

import numpy as np
import math as m
import pandas as pd
import xlrd
import xlwt

constants = {
    '90': 90 * 3600,
    '180': 180 * 3600,
    '270': 270 * 3600,
    '360': 360 * 3600,
    'ro': int(round(180 / np.pi * 3600)),
    'sec_r': np.pi / (180 * 3600)
}


def grad_r_sek(g, m=0, s=0):
    """ Переводит значение угла из градусов в секунды.
        На вход принимает значение угла в формате (g, m, s).
        Обязательный аргумент: g - градусы.
        Два не обязательных аргумента: m - минуты; s - секунды (по умолчанию
        равны 0)
        Возвращает значение угла в секундах.
    """
    s = g * 3600 + m * 60 + s
    return s


def sek_r_grad(s):
    """ Обратная функция для grad_r_sek.
        На вход принимает значение угла в секундах.
        Возвращает список в формате [g, m , s]
        g - градусы; m - минуты; s - секунды.
        Значения секунд округляються по правилам математики до 4 разряда
    """
    A = []
    A.append(int(s // 3600))
    A.append(int(m.fmod(s, 3600) // 60))
    A.append(float('{:.4f}'.format(m.fmod(m.fmod(s, 3600), 60))))
    return A


def gradVrad(angele_in_seconds):
    """ Принимает на вход значение угла в угловых секундах возвращает
        значение угла в радианах
    """
    angle_in_radians = angele_in_seconds * np.pi / (180 * 3600)
    return float('{:.9f}'.format(angle_in_radians))


def radVgrad(angle_in_radians):
    """ Принимает на вход значение угла в радианах возвращает
        значение угла в угловых секундах
    """
    angele_in_seconds = (angle_in_radians * 180 * 3600) / np.pi
    return float('{:.4f}'.format(angele_in_seconds))


def data_excel(path, i=0):
    """ Принимает на вход полный путь к файлу Excel и индекс листа
        (по умолчанию индекс = 0).
        Возврвщает содержимое листа в виде списка
    """
    rb = xlrd.open_workbook(path)
    sheet = rb.sheet_by_index(i)
    return [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

A = [[]] * 16

A[0] = ['№ точки', 'Измеренные углы (левые)', 'Дирекционные углы α', 'Длины линий (м)', 'cos α', 'sin α', 'ΔX (м)',
      'ΔY (м)', 'vΔX', 'vΔY', 'X (м)', 'Y (м)']
i = 1
while i < len(A):
    A[i] = [None] * 12
    i += 1


