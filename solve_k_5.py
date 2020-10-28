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
    'sec_r':np.pi / (180 * 3600),
    'n': 15,
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


def grad_r_sek_list(a=list):
    """ Переводит значение угла из градусов в секунды.
        На вход принимает значение угла в формате (g, m, s).
        Обязательный аргумент: g - градусы.
        Два не обязательных аргумента: m - минуты; s - секунды (по умолчанию
        равны 0)
        Возвращает значение угла в секундах.
    """
    s = a[0] * 3600 + a[1] * 60 + a[2]
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


def pars(str):
    s = str.replace("'", 'º')
    a = s.split('º')
    A = []
    for i in range(3):
        A.append(int(a[i].lstrip()))

    return A[0], A[1], A[2]


def sum_list(a=list, i=1, k=0):
    S = 0
    step = i + 1
    while i < len(a):
        S += a[i][k]
        i += step
    return S


def polygonometry(A=list, B=dict):
    """ Принимает на вход список сформированный в ручную в Excel и словарь с константами.
        Возвращает посчитанную, но не уравненную полигонометрию.
    """
    A[0].append('cosα')
    A[0].append('sinα')
    A[0].append('ΔX')
    A[0].append('ΔY')
    A[0].append('vΔX')
    A[0].append('vΔY')
    A[0].append('X(м)')
    A[0].append('Y(м)')

    for i in A:
        if i[0][0] == 'α' or i[0][0] == 'β':
            i[1] = i[2] * 3600 + i[3] * 60 + i[4]

    for i in range(len(A)):
        if A[i][0][0] == 'α' and A[i][1] == '':
            A[i][1] = A[i - 4][1] + A[i - 2][1] - B['180']
            K = sek_r_grad(A[i][1])
            # print(A[0], A[1], A[2])
            A[i][2] = K[0]
            A[i][3] = K[1]
            A[i][4] = K[2]

    c = 0
    for i in A:
        if i[0][0] == 'S':
            i.append(float(np.cos(gradVrad(A[c - 1][1]))))
            i.append(float(np.sin(gradVrad(A[c - 1][1]))))
            # print(i[1], i[5], i[6])
            i.append(float(i[1] * i[5]))
            i.append(float(i[1] * i[6]))
            c += 1
        else:
            c += 1

    X = A[1][1]
    Y = A[2][1]
    for i in A:
        if i[0][0] == 'S':
            X = X + i[7]
            Y = Y + i[8]
            i.append(X)
            i.append(Y)

    b = []
    i = 0
    while i < len(A):
        if len(A[i]) == 5 and A[i][0] == A[i][1] == A[i][2] == A[i][3] == A[i][4]:
            b.append(A[i])
            del A[i]
        else:
            i += 1
    # print(b)
    return A


def corrrections(a, d, c):
    i = 1
    k = c['n'] - 2
    j = 3
    S = 0
    while j < 43:
        a[j][1] = round(((6 * (k - 2 * (i - 1))) / ((k + 1) * (k + 2))) * 10, 1)
        S += a[j][1]
        # print(a[j][1])
        i += 1
        j += 3
    Svb = round(S, 1)
    # print('S =', round(S, 3))

    i = 4
    S = 0
    while i < 46:
        a[i][3] = round(- ((d['t'] / d['L']) * a[i + 1][3]) * 10 ** 3, 1)
        S += a[i][3]
        # print(a[i][3])
        i += 3
    Sv = S

    a[1][2] = round(d['Δα'], 1)
    a[5][2] = a[1][2] + a[3][1]
    # print(a[1][2], a[5][2], sep='\n')
    i = 8
    S = a[1][2] + a[5][2]
    while i < 46:
        a[i][2] = round(a[i - 3][2] + a[i - 2][1], 1)
        S += a[i][2]
        # print(a[i][2])
        i += 3
    Sa = S

    a[2][7] = -(a[2][6] * a[1][2]) / c['ro']
    a[2][8] = (a[2][5] * a[1][2]) / c['ro']
    S = a[2][7]
    S1 = a[2][8]

    i = 5
    while i < 46:
        a[i][7] = round(a[i - 1][3] * np.cos(gradVrad(grad_r_sek_list(pars(a[i + 1][2])))) -
                        (a[i][6] * a[i][2]) / c['ro'], 1)
        a[i][8] = round(a[i - 1][3] * np.sin(gradVrad(grad_r_sek_list(pars(a[i + 1][2])))) +
                        (a[i][5] * a[i][2]) / c['ro'], 1)
        S += a[i][7]
        S1 += a[i][8]
        # print(i, ':', a[i][7], '---', a[i][8])
        i += 3
    Svx = S
    Svy = S1

    S = 0
    i = 4
    while i < 46:
        S += grad_r_sek_list(pars(a[i][1]))
        i += 3
        # print(a[i][2])
    Sb = sek_r_grad(S)


    return a, Svb, Sa, Sv, Svx, Svy, Sb


raw_data = (
    r'F:\Документы\МИИГАиК\МИИГиКА(четвёртый курс)\Контрольные\Контрольная №5\К_Контрольной_5.xlsx'
)
source_data = data_excel(raw_data)
# C = polygonometry(source_data, constants)
# df = pd.DataFrame(C)
# df.to_excel("Чистенько.xlsx", sheet_name='Лист1')

calculated_data = data_excel(raw_data, 2)

discrepancies = {
    'fx': calculated_data[45][9] - calculated_data[46][9],
    'fy': calculated_data[45][10] - calculated_data[46][10],
    'fs': np.sqrt((calculated_data[45][9] - calculated_data[46][9]) ** 2 +
                  (calculated_data[45][10] - calculated_data[46][10]) ** 2),
    'S': sum_list(calculated_data, 2, 3),
    'L': sum_list(calculated_data, 2, 3) - (calculated_data[2][3] + calculated_data[44][3]),
    'Δx': sum_list(calculated_data, 2, 5),
    'Δy': sum_list(calculated_data, 2, 6),
}
discrepancies['t'] = ((discrepancies['fy'] * discrepancies['Δy'] + discrepancies['fx'] * discrepancies['Δx']) /
                      discrepancies['L'])
discrepancies['u'] = ((discrepancies['fy'] * discrepancies['Δx'] - discrepancies['fx'] * discrepancies['Δy']) /
                      discrepancies['L'])
discrepancies['mu1'] = discrepancies['L'] * 10 ** 3 / 45000
discrepancies['mu2'] = (4 * discrepancies['L'] * 10 ** 3) / constants['ro'] * np.sqrt((constants['n'] + 1.5) / 3)
discrepancies['mu3'] = (8 * discrepancies['L'] * 10 ** 3) / constants['ro']
discrepancies['mu'] = np.sqrt(discrepancies['mu1'] ** 2 + discrepancies['mu2'] ** 2 + discrepancies['mu3'] ** 2)
discrepancies['U3'] = (discrepancies['u'] * 10 ** 3 * discrepancies['mu3'] ** 2) / discrepancies['mu'] ** 2
discrepancies['Δα'] = - (discrepancies['U3'] * constants['ro']) / (discrepancies['L'] * 10 ** 3)
discrepancies['ω'] = - (((discrepancies['u'] * 10 ** 3 - discrepancies['U3']) * constants['ro']) /
                        (discrepancies['L'] * 10 ** 3))

a, Svb, Sa, Sv, Svx, Svy, Sb = corrrections(calculated_data, discrepancies, constants)

discrepancies['∑β'] = Sb
discrepancies['∑υβ'] = Svb
discrepancies['∑υα'] = Sa
discrepancies['∑υs'] = Sv
discrepancies['∑υΔx'] = Svx
discrepancies['∑υΔy'] = Svy

# df = pd.DataFrame(a)
# df.to_excel("Итоговая.xlsx", sheet_name='Лист1')

#a[i][4]


# discrepancies_list = discrepancies.items()
# df = pd.DataFrame(discrepancies_list)
# df.to_excel("Вычисления.xlsx", sheet_name='Лист1')



for i in discrepancies:
    print('{:4} : {}'.format(i, discrepancies[i]))

    








