from matplotlib import pyplot
from matplotlib import patches

def find_missing_arrays(data_new,data_old):
    for array_new in data_new:
        arr_found = False
        for array_old in data_old:
            if array_new[2] == array_old[2]:
                arr_found = True
        if arr_found == False:
            #print('Array',array_new[2],'appears in New data, but not found in Old data. Skipping.')
            print(array_new[2],'появился в новых данных')
    for array_old in data_old:
        arr_found = False
        for array_new in data_new:
            if array_new[2] == array_old[2]:
                arr_found = True
        if arr_found == False:
            #print('Array',array_old[2],'not found in New data, but was present in Old data. Skipping.')
            print(array_old[2],'отсутствует в новых данных')

def produce_delta_graph_01(data_new,data_old):
    find_missing_arrays(data_new,data_old)
    found_arrays = []
    for n,array_new in enumerate(data_new):
        for o,array_old in enumerate(data_old):
            if array_new[2] == array_old[2]:
                found_arrays.append([array_new,array_old])

    save_graph(found_arrays, 'Graph_04')

def produce_delta_graph_02(data_new,data_old):
    found_arrays = []
    for n,array_new in enumerate(data_new):
        for o,array_old in enumerate(data_old):
            if array_new[2] == array_old[2]:
                found_arrays.append([array_new,array_old])
    risky_arrays = []
    for found_array in found_arrays:
        if found_array[0][0] > 50 and found_array[0][1] > 100:
            risky_arrays.append(found_array)
    save_graph(risky_arrays, 'Graph_05')


def save_graph(draw_data, graph_name):
    fig = pyplot.figure(figsize=(16,8))
    #ax = fig.add_subplot(1,1,1)
    ax = pyplot.subplot2grid((3,1), (0,0), rowspan=2, colspan=1)
    pyplot.title("Total Demand / Utilization %", fontsize = 12)

    dynamic_txt_offset = 0.4
    dynamic_txt_offset_2 = 0.4

    if len(draw_data) >= 80:
        dynamic_font_size = 4

    elif len(draw_data) < 20:
        dynamic_font_size = 8
        dynamic_txt_offset = 0.0
    elif len(draw_data) < 40:
        dynamic_font_size = 7

    elif len(draw_data) < 60:
        dynamic_font_size = 6

    elif len(draw_data) < 80:
        dynamic_font_size = 5


    for y in range(1,7):
        if y == 5 or y == 10:
            pyplot.axhline(y*20,color='#FF2F1A',linewidth=0.7 ,alpha=1, zorder=4) # '#B6B6B6'
        else:
            pyplot.axhline(y*20,color='#B6B6B6',linewidth=0.7,linestyle='--' ,alpha=0.7, zorder=1) # '#B6B6B6'

    for a,array in enumerate(draw_data):
        # Array name
        pyplot.text(a - dynamic_txt_offset, 5, array[0][2], fontsize=dynamic_font_size, rotation=90, rotation_mode='anchor', horizontalalignment='left', verticalalignment='top')
        # Demand
        ax.bar(a + 0.3, array[0][1], 0.4, bottom=0.001, color='#C2C2C2', edgecolor='black', linewidth='0.3', alpha=0.5, zorder=2) # '#7FF9E2'
        # Utilization
        ax.bar(a + 0.3, array[0][0], 0.4, bottom=0.001, color='#0D5265', edgecolor='black', linewidth='0.3', alpha=1, zorder=3) # '#0D5265'

    ax.set_yscale('linear')

    pyplot.tick_params(
    axis='x',          # changes apply to the x-axis
    which='both',      # both major and minor ticks are affected
    bottom=False,      # ticks along the bottom edge are off
    top=False,         # ticks along the top edge are off
    labelbottom=False)

    #ax.set_xlabel('x')
    ax.set_ylabel('Percent %')

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    legend_elements = [patches.Patch(facecolor='#0D5265', edgecolor='black', linewidth='0.3', label='Utilization'),patches.Patch(facecolor='#C2C2C2', edgecolor='black', linewidth='0.3', label='Demand', alpha=0.5)]
    ax.legend(handles=legend_elements, bbox_to_anchor=(1.122, 1), ncol=1, frameon=False) # loc='center right'

    ax_2 = pyplot.subplot2grid((3,1), (2,0), rowspan=1, colspan=1)
    pyplot.title("Trend for Demand / Utilization, delta %", fontsize = 12)

    pyplot.tick_params(
    axis='x',          # changes apply to the x-axis
    which='both',      # both major and minor ticks are affected
    bottom=False,      # ticks along the bottom edge are off
    top=False,         # ticks along the top edge are off
    labelbottom=False)

    ax_2.spines['top'].set_visible(False)
    ax_2.spines['bottom'].set_visible(False)
    ax_2.spines['right'].set_visible(False)
    ax_2.spines['left'].set_visible(False)

    pyplot.axhline(0,color='black',linewidth=0.8 ,alpha=1, zorder=1)

    for a,array in enumerate(draw_data):
        # Array name
        ##pyplot.text(a - 0.1, 0.1, array[0][2], fontsize=8, rotation=90, rotation_mode='anchor', horizontalalignment='left', verticalalignment='top')
        # Util delta bar
        ax_2.bar(a + 0.3, array[0][0]-array[1][0], 0.2, bottom=0.001, color='#0D5265', edgecolor='black', linewidth='0.3', alpha=1, zorder=1) # '#0D5265'
        # Demand delta bar
        ax_2.bar(a + 0.5, array[0][1]-array[1][1], 0.2, bottom=0.001, color='#C2C2C2', edgecolor='black', linewidth='0.3', alpha=0.5, zorder=1) # '#7FF9E2'
        # Util delta text
        if len(draw_data) > 10:
            upr_txt_ofst = 0.12
            lwr_txt_ofst = upr_txt_ofst/3
            if (array[0][0]-array[1][0]) < 0:
                pyplot.text(a + dynamic_txt_offset_2, array[0][0]-array[1][0]-lwr_txt_ofst, round(abs(array[0][0]-array[1][0]),2), fontsize=dynamic_font_size, rotation=90, rotation_mode='anchor', horizontalalignment='right', verticalalignment='bottom', color='grey')
            elif (array[0][0]-array[1][0]) > 0:
                pyplot.text(a + dynamic_txt_offset_2, array[0][0]-array[1][0]+upr_txt_ofst, round(array[0][0]-array[1][0],2), fontsize=dynamic_font_size, rotation=90, rotation_mode='anchor', horizontalalignment='left', verticalalignment='bottom', color='grey')
            # Demand delta text
            if (array[0][1]-array[1][1]) < 0:
                pyplot.text(a + dynamic_txt_offset_2, array[0][1]-array[1][1]-lwr_txt_ofst, round(abs(array[0][1]-array[1][1]),2), fontsize=dynamic_font_size, rotation=90, rotation_mode='anchor', horizontalalignment='right', verticalalignment='top', color='grey')
            elif (array[0][1]-array[1][1]) > 0:
                pyplot.text(a + dynamic_txt_offset_2, array[0][1]-array[1][1]+upr_txt_ofst, round(array[0][1]-array[1][1],2), fontsize=dynamic_font_size, rotation=90, rotation_mode='anchor', horizontalalignment='left', verticalalignment='top', color='grey')
        else:
            upr_txt_ofst = 0.03
            lwr_txt_ofst = upr_txt_ofst/3
            if (array[0][0]-array[1][0]) < 0:
                pyplot.text(a + dynamic_txt_offset_2, array[0][0]-array[1][0]-lwr_txt_ofst, round(abs(array[0][0]-array[1][0]),2), fontsize=dynamic_font_size, rotation=0, rotation_mode='anchor', horizontalalignment='right', verticalalignment='top', color='grey')
            elif (array[0][0]-array[1][0]) > 0:
                pyplot.text(a + dynamic_txt_offset_2, array[0][0]-array[1][0]+upr_txt_ofst, round(array[0][0]-array[1][0],2), fontsize=dynamic_font_size, rotation=0, rotation_mode='anchor', horizontalalignment='right', verticalalignment='bottom', color='grey')
            # Demand delta text
            if (array[0][1]-array[1][1]) < 0:
                pyplot.text(a + dynamic_txt_offset_2, array[0][1]-array[1][1]-lwr_txt_ofst, round(abs(array[0][1]-array[1][1]),2), fontsize=dynamic_font_size, rotation=0, rotation_mode='anchor', horizontalalignment='left', verticalalignment='top', color='grey')
            elif (array[0][1]-array[1][1]) > 0:
                pyplot.text(a + dynamic_txt_offset_2, array[0][1]-array[1][1]+upr_txt_ofst, round(array[0][1]-array[1][1],2), fontsize=dynamic_font_size, rotation=0, rotation_mode='anchor', horizontalalignment='left', verticalalignment='bottom', color='grey')


    legend_elements = [patches.Patch(facecolor='#0D5265', edgecolor='black', linewidth='0.3', label='Utilization'),patches.Patch(facecolor='#C2C2C2', edgecolor='black', linewidth='0.3', label='Demand', alpha=0.5)]
    ax_2.legend(handles=legend_elements, bbox_to_anchor=(1.122, 1), ncol=1, frameon=False) # loc='center right'

    ax_2.set_yscale('linear')


    pyplot.savefig(str(graph_name)+'.png', dpi = 300)
