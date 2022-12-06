import argparse
from Graph_01_3Par_096 import produce_graph_01_variants
from Graph_02_3Par_096 import produce_graph_02_variants
from Graph_03_3Par_01 import produce_graph_03
from tPar_capacity_data_read_utils import get_arrays_xlsx,get_arrays_xlsx_full
from classify_3Par_arrays_capacity import combine_arrays_data

if __name__ == '__main__':

    # Args
    parser = argparse.ArgumentParser(description='Produces multiple variants of 3Par utilization/overprovisioning graphs 1,2 using input xlsx tables')
    parser.add_argument('-st1', type=str, help='Excel sheet for current data', required=True)
    parser.add_argument('-st2', type=str, help='Excel sheet for comparison')
    parser.add_argument('-iter', type=int, help='Number of iterations (increasing labels distance each iteration), default: 1')
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

    print('Processing data:',st1,'  Compare with:',st2,'  Graph iterations:',iter_num)

    # Start real work
    st1_points_xy = get_arrays_xlsx(st1)

    print('\nProducing Graph 01\n')
    # produce_graph_01_variants(st1_points_xy,iter_num)
    print('\nProducing Graph 02\n')
    # produce_graph_02_variants(st1_points_xy,iter_num)

    if st2 != None:
        print('\nProducing Graph 03\n')
        st2_points_xy = get_arrays_xlsx(st2)
        produce_graph_03(st1_points_xy, st2_points_xy)
        print('\nGet additional Capacity data from input Excel sheets\n')
        st1_arrays = get_arrays_xlsx_full(st1)
        st2_arrays = get_arrays_xlsx_full(st2)
        print('\nClasifying arrays data and producing Excel tables:\n')
        arrays_data = combine_arrays_data(st1,st1_arrays,st2,st2_arrays)

    pressenter = input('\nPress Enter to exit.')
