def sortMasiv(masiv):
    for i in range(len(masiv)):
        masiv[i] = list(masiv[i])

    for i in range(len(masiv)):
        if masiv[i][3] == '1x25':
            masiv[i][3] = 1
        elif masiv[i][3] == '1x35':
            masiv[i][3] = 2
        elif masiv[i][3] == '2x10':
            masiv[i][3] = 3
        elif masiv[i][3] == '2x15':
            masiv[i][3] = 4
        elif masiv[i][3] == '2x20x40':
            masiv[i][3] = 5
        elif masiv[i][3] == '2x50':
            masiv[i][3] = 6
        elif masiv[i][3] == 'Затравочное':
            masiv[i][3] = 7
        elif masiv[i][3] == 'Плавка':
            masiv[i][3] = 8
        elif masiv[i][3] == 'Л/С':
            masiv[i][3] = 9

        if masiv[i][4] == 'V - 4,0':
            masiv[i][4] = 1
        elif masiv[i][4] == 'V - 5,0':
            masiv[i][4] = 2
        elif masiv[i][4] == 'V - 6,0':
            masiv[i][4] = 3
        elif masiv[i][4] == 'V - 7,0':
            masiv[i][4] = 4

    masiv = sorted(masiv, key=lambda a: (a[3], a[4]))

    for i in range(len(masiv)):
        if masiv[i][3] == 1:
            masiv[i][3] = '1x25'
        elif masiv[i][3] == 2:
            masiv[i][3] = '1x35'
        elif masiv[i][3] == 3:
            masiv[i][3] = '2x10'
        elif masiv[i][3] == 4:
            masiv[i][3] = '2x15'
        elif masiv[i][3] == 5:
            masiv[i][3] = '2x20x40'
        elif masiv[i][3] == 6:
            masiv[i][3] = '2x50'
        elif masiv[i][3] == 7:
            masiv[i][3] = 'Затравочное'
        elif masiv[i][3] == 8:
            masiv[i][3] = 'Плавка'
        elif masiv[i][3] == 9:
            masiv[i][3] = 'Л/С'

        if masiv[i][4] == 1:
            masiv[i][4] = 'V - 4,0'
        elif masiv[i][4] == 2:
            masiv[i][4] = 'V - 5,0'
        elif masiv[i][4] == 3:
            masiv[i][4] = 'V - 6,0'
        elif masiv[i][4] == 4:
            masiv[i][4] = 'V - 7,0'
    return masiv
