import argparse
from tPar_capacity_data_read_utils import get_arrays_xlsx_full
from excel_output_utils import write_collected_data_to_worksheet, save_workbook


def combine_arrays_data(st1,st1_arrays,st2,st2_arrays):

    found_arrays = []
    for st1_array in st1_arrays:
        for st2_array in st2_arrays:
            if st1_array[0] == st2_array[0]:
                # Data structure: 0 Array name 1 X% 2 Y% 3 Total Raw 4 Raw 5 old X% 6 old Y% 7 old Total Raw 8 old Raw
                found_arrays.append([st1_array[0],st1_array[1],st1_array[2],st1_array[3],st1_array[4],st2_array[1],st2_array[2],st2_array[3],st2_array[4]])

    all_systems = []
    risk_systems = []
    rest_systems = []
    for array in found_arrays:
        array_name = array[0]
        array_x = array[1]
        array_y = array[2]
        array_delta_x = round(array[1]-array[5],2)
        array_delta_y = round(array[2]-array[6],2)
        array_delta_x_tb = round(array[4]-array[8],2)
        overprov_tb = array[3] * array[2]/100
        array_delta_y_tb = round(overprov_tb * (array[2] - array[6])/100,2)

        all_systems.append([array_name,array_x,array_y,array_delta_x,array_delta_y,array_delta_x_tb,array_delta_y_tb])

        if array[1] > 75 and array[2] > 100:
            risk_systems.append([[array_name,array_x,array_y,array_delta_x,array_delta_y,array_delta_x_tb,array_delta_y_tb]])
        else:
            rest_systems.append([[array_name,array_x,array_y,array_delta_x,array_delta_y,array_delta_x_tb,array_delta_y_tb]])

    systems_headers = ['System Name','X%','Y%','Delta X%','Delta Y%','Delta X (Tb)','Delta Y (Tb)']
    write_collected_data_to_worksheet('Risk',systems_headers,risk_systems)
    write_collected_data_to_worksheet('Rest',systems_headers,rest_systems)



    saved = False
    output_filename = st2.replace('.xlsx','') + '_' + st1.replace('.xlsx','') + '_classified'
    while not saved:
        try:
            save_workbook(output_filename)
            saved = True
        except:
            message = input('File:' + output_filename + '.xlsx\" is busy. Press Enter to retry')

    return all_systems


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Produces Excel workbook with two worksheets: 1.Systems in risk zone   2.Rest systems  ')
    parser.add_argument('-st1', type=str, help='Excel sheet for current data', required=True)
    parser.add_argument('-st2', type=str, help='Excel sheet for comparison data', required=True)
    
    args = parser.parse_args()

    st1 = args.st1
    st2 = args.st2

    st1_arrays = get_arrays_xlsx_full(st1)
    st2_arrays = get_arrays_xlsx_full(st2)

    combine_arrays_data(st1,st1_arrays,st2,st2_arrays)
