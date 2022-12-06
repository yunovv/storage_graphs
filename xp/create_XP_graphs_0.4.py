import argparse
import openpyxl
from Graph_kv_XP_002 import produce_kv_graph
from XP_capacity_data_read_utils import get_arrays_xlsx,get_arrays_xlsx_compr,get_arrays_xlsx_gr03
from Graph_delta_XP_001 import produce_delta_graph_01,produce_delta_graph_02


if __name__ == '__main__':

    # Args
    parser = argparse.ArgumentParser(description='Produces multiple variants of XP utilization/overprovisioning graphs using input xlsx tables')
    parser.add_argument('-st1', type=str, help='Excel sheet for current data', required=True)
    parser.add_argument('-st2', type=str, help='Excel sheet for comparison')
    parser.add_argument('-iter', type=int, help='Number of iterations (increasing labels distance each iteration), default: 1')
    parser.add_argument('-skip', type=int, help='Number of Graphs to skip, default: 0')
    args = parser.parse_args()

    # main Excel sheet
    if args.st1 == None or '.xlsx' not in args.st1 or len(args.st1) < 6:
        st1 = ''
        while '.xlsx' not in st1 or len(st1) < 6:
            st1 = input('Please specify correct Excel sheet name for current data(*.xlsx): ')
    else:
        st1 = args.st1

    # comparison Excel sheet
    if args.st1 == None:
        if args.st2 != None and '.xlsx' in args.st2  and len(args.st2) > 5:
            st2 = args.st2
        else:
            st2 = None
            while st2 == None or (st2 != '' and '.xlsx' not in st2) or (st2 != '' and len(st2) < 6):
                st2 = input('Please specify Excel sheet name for comparison(*.xlsx) or press Enter to skip: ')
            if st2 == '':
                st2 = None
    else:
        if args.st2 != None and '.xlsx' in args.st2  and len(args.st2) > 5:
            st2 = args.st2
        else:
            st2 = None

    # number of iterations
    if args.iter == None or args.iter <= 0:
        iter_num = 1
    else:
        iter_num = args.iter

    if args.skip == None or args.skip <=0:
        skip = 0
    else:
        skip = args.skip

    print('Processing data:',st1,'  Compare with:',st2,'  Graph iterations:',iter_num)

    # Start real work
    print('\nProducing output Tables\n')

    if skip < 1:
        print('\nProducing Graph 01\n')
        st1_points_xy = get_arrays_xlsx(st1)
        graph_01_orange_area = [50,100,100]
        graph_01_labels = ['LOGICAL OVERPROV VS UTILIZATION','Используемые ресурсы, %','Потребность, %']
        graph_01_name = 'Graph_01'
        produce_kv_graph(st1_points_xy,iter_num,graph_01_labels,graph_01_name)
    else:
        print('\nSkipping Graph 01\n')


    if skip < 2:
        print('\nProducing Graph 02\n')
        st1_points_xy_compr = get_arrays_xlsx_compr(st1)
        graph_02_orange_area = [50,100,100]
        graph_02_labels = ['LOGICAL OVERPROV VS UTILIZATION (COMPR > 1)','Используемые ресурсы, %','Потребность, %']
        graph_02_name = 'Graph_02'
        produce_kv_graph(st1_points_xy_compr,iter_num,graph_02_labels,graph_02_name)
    else:
        print('\nSkipping Graph 02\n')

    if skip < 3:
        print('\nProducing Graph 03\n')
        st1_points_xy_gr03 = get_arrays_xlsx_gr03(st1)
        graph_03_orange_area = [50,100,100]
        graph_03_labels = ['PHYSICAL OVERPROV VS UTILIZATION (COMPR > 1 & HIGH POOL OVERPROV/UTILIZATION OR EXPANSION>COMPR)','Использование ресурсов, %','Потребность, %']
        graph_03_name = 'Graph_03'
        produce_kv_graph(st1_points_xy_gr03,iter_num,graph_03_labels,graph_03_name)
    else:
        print('\nSkipping Graph 03\n')

    if st2 != None:
        if skip < 4:
            print('\nProducing Graph 04\n')
            st1_points_xy = get_arrays_xlsx(st1)
            st2_points_xy = get_arrays_xlsx(st2)
            produce_delta_graph_01(st1_points_xy, st2_points_xy)
            produce_delta_graph_02(st1_points_xy, st2_points_xy)
            # print('\nGet additional Capacity data from input Excel sheets\n')
            # st1_arrays = get_arrays_xlsx_full(st1)
            # st2_arrays = get_arrays_xlsx_full(st2)
            # print('\nClasifying arrays data and producing Excel tables:\n')
            # arrays_data = combine_arrays_data(st1,st1_arrays,st2,st2_arrays)
            # print('\nProducing Graph 04\n')
            # produce_graph_05(arrays_data)
        else:
            print('\nSkipping Graph 04\n')
    #
    # pressenter = input('\nPress Enter to exit.')
