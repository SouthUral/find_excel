test_dict_1 = {'А-1-1': 
                {'date_1': []
                ,'date_2': ['T-650']
                ,'date_3': ['T-630', 'T-530', 'T-730', 'T-830', 'T-700', 'T-900']
                ,'date_4': ['T-401', 'T-501']}}

def matrix_production(dict_matrix):
    new_matrix = []
    for value in dict_matrix.values():
        new_matrix.append(value)
    return new_matrix


def transposed(matrix):
    count = 0
    new_matrix = []
    while True:
        new_string = []
        for string in matrix:
            if (count+1) > len(string):
                new_string.append('0')
            else:
                new_string.append(string[count])
        if new_string == ['0','0','0','0']:
            break
        else:
            new_matrix.append(list(new_string))
            count += 1
    return new_matrix


def data_to_write(key:str, value:list):
    pre_matrix = matrix_production(value)
    pre_matrix_1 = transposed(pre_matrix)
    for string in pre_matrix_1:
        string.insert(0, key)
    return pre_matrix_1
    

