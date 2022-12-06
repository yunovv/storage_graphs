from matplotlib import pyplot
from matplotlib import patches

def produce_graph_03(data_old,data_new):
    found_arrays = []
    for n,array_new in enumerate(data_new):
        for o,array_old in enumerate(data_old):
            if array_new[2] == array_old[2]:
                found_arrays.append([array_new,array_old])

    save_graph(found_arrays, 'Graph_03')


def save_graph(draw_data, graph_name):
    fig = pyplot.figure(figsize=(16,8))
    #ax = fig.add_subplot(1,1,1)
    ax = pyplot.subplot2grid((3,1), (0,0), rowspan=2, colspan=1)
    pyplot.title("Total Demand / Utilization %", fontsize = 12)

    for y in range(1,7):
        if y == 5 or y == 10:
            pyplot.axhline(y*20,color='#FF2F1A',linewidth=0.7 ,alpha=1, zorder=4) # '#B6B6B6'
        else:
            pyplot.axhline(y*20,color='#B6B6B6',linewidth=0.7,linestyle='--' ,alpha=0.7, zorder=1) # '#B6B6B6'

    for a,array in enumerate(draw_data):
        ax.bar(a + 0.3, array[0][1], 0.4, bottom=0.001, color='#C2C2C2', edgecolor='black', linewidth='0.3', alpha=0.5, zorder=2) # '#7FF9E2'
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

    legend_elements = [patches.Patch(facecolor='#0D5265', edgecolor='black', linewidth='0.3', label='Utilization'),patches.Patch(facecolor='#C2C2C2', edgecolor='black', linewidth='0.3', label='Demand')]
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
        ax_2.bar(a + 0.3, array[0][0]-array[1][0], 0.2, bottom=0.001, color='#0D5265', edgecolor='black', linewidth='0.3', alpha=1, zorder=1) # '#0D5265'
        ax_2.bar(a + 0.5, array[0][1]-array[1][1], 0.2, bottom=0.001, color='#7FF9E2', edgecolor='black', linewidth='0.3', alpha=1, zorder=1) # '#7FF9E2'
        pyplot.text(a - 0.1, -0.1, array[0][2], fontsize=8, rotation=90, rotation_mode='anchor', horizontalalignment='right', verticalalignment='top') #

    legend_elements = [patches.Patch(facecolor='#0D5265', edgecolor='black', linewidth='0.3', label='Utilization'),patches.Patch(facecolor='#7FF9E2', edgecolor='black', linewidth='0.3', label='Demand')]
    ax_2.legend(handles=legend_elements, bbox_to_anchor=(1.122, 1), ncol=1, frameon=False) # loc='center right'

    ax_2.set_yscale('linear')


    pyplot.savefig(str(graph_name)+'.png', dpi = 300)
