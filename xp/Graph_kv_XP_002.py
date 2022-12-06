import openpyxl
from matplotlib import pyplot,patches
import math
import sys
from shapely.geometry import LineString
import colour
import numpy as np
import concurrent.futures
#import multiprocessing
##from tqdm.auto import tqdm
import datetime


# Sort points based on distance to borders
def make_sequence_brd_dist():
    brdr_dist_coef = []
    for point in points_xy:
        tmp_brdr_coef = 400
        if abs(point[0] - xmax)*2 < tmp_brdr_coef:
            tmp_brdr_coef = abs(point[0] - xmax)*2
        if abs(point[0] - 0)*2 < tmp_brdr_coef:
            tmp_brdr_coef = abs(point[0] - 0)*2
        if abs(point[1]  - 0) < tmp_brdr_coef:
            tmp_brdr_coef = abs(point[1]  - 0)
        if abs(point[1]  - ymax) < tmp_brdr_coef:
            tmp_brdr_coef = abs(point[1]  - ymax)
        brdr_dist_coef.append(tmp_brdr_coef)
    #print(brdr_dist_coef)
    points_with_dist = []
    for i,point in enumerate(points_xy):
        points_with_dist.append([*point,brdr_dist_coef[i]])
    points_with_dist_sorted = sorted(points_with_dist, key = lambda x: x[3])
    points_with_dist_sorted_wo_coef = []
    for point in points_with_dist_sorted:
        point.pop(3)
        points_with_dist_sorted_wo_coef.append(point)
    return points_with_dist_sorted_wo_coef
#print(points_with_dist_sorted_wo_coef)



# Sort points based on their biggest angle
def make_sequence_max_angle():
    points_with_max_angle = []
    for point in points_xy:
        angles_data = determine_label_preferred_angle(point)
        angles_data_sorted = sorted(angles_data, key = lambda x: x[6], reverse = True)
        points_with_max_angle.append(point+[angles_data_sorted[0][5]])

    points_with_max_angle_sorted = sorted(points_with_max_angle, key = lambda x: x[3])
    points_with_max_angle_sorted_wo_angle_width = []
    for point in points_with_max_angle_sorted:
        point.pop(3)
        points_with_max_angle_sorted_wo_angle_width.append(point)
    return points_with_max_angle_sorted_wo_angle_width



# Apply colours based on point location
def colour_points(variant):
    orange_0 = colour.Color('#FFB84C')
    orange_colors = list(orange_0.range_to(colour.Color('#E01900'),100))
    orange_colors_str = []
    for orange_color in orange_colors:
        orange_colors_str.append(str(orange_color).strip('<Color ').strip('>'))
    #print(orange_colors_str)
    #print(orange_colors)

    p_colors = []
    for point in variant:
        #p_colors.append('grey')
        if point[0] <= xmax/2 or point[1] <= ymax/2:
            ##p_colors.append('green')
            p_colors.append('#01A982')
        elif point[0] <= xmax and point[1] <= ymax:
            ##p_colors.append('orange')
            p_colors.append(orange_colors_str[int(point[0]-xmax/2 + (point[1]-ymax/2)/2)-1])
    return p_colors


def get_angle_precise(a, b, c):
    ang = math.degrees(math.atan2(c[1]-b[1], c[0]-b[0]) - math.atan2(a[1]-b[1], a[0]-b[0]))
    return ang + 360 if ang < 0 else ang



# Form angles to check list for range of min max angles
def get_angles_between_2(min,max):
    angles = []
    detail_level = 3
    angle_offset = 10
    if max < min:
        max += 360
    for angle in range(int(((max-angle_offset*2) - min)/detail_level)+1):
        angles.append(min + angle_offset + angle*detail_level)
    if len(angles) == 0:
        angles.append((min + max)/2)
    return angles

def get_angles_between(min,max):
    angles_centrified = []
    detail_level = 3
    angle_offset = 10
    if max <= min:
        max += 360
    number_of_angles = int(((max-angle_offset*2) - min)/detail_level)+1
    #print('number of angles',number_of_angles)
    middle_angle = (min + max)/2
    angles_centrified.append(middle_angle)
    for iter in range(1,int(number_of_angles/2)+1):
        angles_centrified.append(middle_angle+iter*detail_level)
        angles_centrified.append(middle_angle-iter*detail_level)

    #print('min',min,'max',max,'angles between',angles_centrified)
    return angles_centrified

def check_lines_cross(line_a_coords,line_b_coords):
    line_a = LineString(line_a_coords)
    line_b = LineString(line_b_coords)
    return line_a.intersects(line_b)

def get_other_points_list(my_point):
    other_points = []
    for point in points_xy:
        if point[2] != my_point[2]:
            other_points.append(point)
    return other_points

def normalize_angle(angle):
    while angle > 360:
        angle -= 360
    return angle

def axis_crossed(point_a,point_b):
    axis_crossed_coef = 0
    if check_lines_cross([point_a,point_b],[[0,ymax/2],[xmax,ymax/2]]) or check_lines_cross([point_a,point_b],[[xmax/2,0],[xmax/2,ymax]]):
        axis_crossed_coef += 10
    return axis_crossed_coef

def get_points_distance(point_a,point_b):
    distance = math.sqrt((point_b[0] - point_a[0])**2 + (point_b[1] - point_a[1])**2)
    return distance

# Debug points angles
# debug_arrows = []
# for point_xy in points_xy:
#     if point_xy[2] == 'Not3par-6sk':
#         for angle in [0,45,90,135,180,225,270,315]:
#             debug_point = [point_xy[0] + 20*get_angle_verthor_coef(angle) * math.sin(math.radians(angle)),point_xy[1] - 20*get_angle_verthor_coef(angle) * math.cos(math.radians(angle))]
#             debug_arrows.append([point_xy[0],point_xy[1],debug_point[0],debug_point[1],angle])
# Debug


#for point in points_xy:
def determine_label_preferred_angle(point):
    ##print('\nChecking point',point[2])
    label_angle_data = []

    tmp_discovered_points = get_points_around(point,scan_radius,5)
    #print('discovered points tmp',tmp_discovered_points)

    # Determine precise angle of point location
    discovered_points_unsorted = []
    discovered_points_unsorted.extend(tmp_discovered_points)
    for disc_point in discovered_points_unsorted:
        #print(disc_point[3])
        disc_point[3] = get_angle_precise([point[0],point[0]-50],[point[0],point[1]],[disc_point[0],disc_point[1]])
        #print(disc_point[3])
    # Sort points using precise angle of point location
    discovered_points = []
    discovered_points = sorted(discovered_points_unsorted, key = lambda x: x[3])


    ##print('discovered points',discovered_points)
    ##if len(discovered_points) != len(other_points):
    ##    print('- - - Alert!: discovered',len(discovered_points),'of',len(other_points))
    ##print('calculating vacant angles')

    label_angle = -1
    for m in range(len(discovered_points)):
        if m == len(discovered_points)-1:
            n = 0
            ##print('last and first',discovered_points[m][2],discovered_points[m][3],discovered_points[n][2],discovered_points[n][3])
            temp_angle = discovered_points[m][3] + (get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]]))/2

            #Debug
            #print('preprocessed angle width =',360 - discovered_points[m][3] + discovered_points[n][3])
            ##print('precise angle width =',get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]]))
            #Debug
            angle_width = get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]])


        else:
            n = m+1
            ##print('current and next',discovered_points[m][2],discovered_points[m][3],discovered_points[n][2],discovered_points[n][3])
            temp_angle = discovered_points[m][3] + (get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]]))/2

            #Debug
            #print('preprocessed angle width =',discovered_points[n][3] - discovered_points[m][3])
            ##print('precise angle width =',get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]]))
            #Debug
            angle_width = get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]])

        print('Angle between',discovered_points[m][2],discovered_points[n][2])
        cos_coef = angle_cos_coef(normalize_angle(temp_angle),angle_width)
        label_angle_data.append([normalize_angle(temp_angle), temp_angle-angle_width/2, temp_angle+angle_width/2, discovered_points[m][2], discovered_points[n][2],angle_width,cos_coef])

        ##print('label angle',label_angle_data)
    # label_angle: recommended angle, angle minimum, angle maximum, first point of angle, second point of angle
    ##print('Angle for label placement:',label_angle_data)
    if len(discovered_points) < 2:
        label_angle_data.append([180, 0, 360, '0', '360', 360, 360])

    return label_angle_data

def angle_cos_coef(angle,angle_width):
    coef = angle_width*(abs(math.cos(math.radians(angle)))**1)
    print(angle,angle_width,coef)
    return coef

def point_line_dist(point,line):
    norm = np.linalg.norm

    p1 = np.array([line[0][0],line[0][1]])
    p2 = np.array([line[1][0],line[1][1]])

    p3 = np.array([point[0], point[1]])
    dist = np.abs(norm(np.cross(p2-p1, p1-p3)))/norm(p2-p1)
    return dist


def label_size(text):
    ls = 0
    for char in text:
        if str.isupper(char):
            ls += 1.3*char_coef
        else:
            ls += 1*char_coef
    return ls

# Dynamic annotation placement delta calculation
#occupied_dots = []
def occupy_points_lines(points):
    occupied_points_lines = []
    for point in points:
        occupied_points_lines.append([[point[0]-p_occ_s*3/20,point[1]+p_occ_s],[point[0]+p_occ_s*3/20,point[1]-p_occ_s]])
        occupied_points_lines.append([[point[0]-p_occ_s*3/20,point[1]-p_occ_s],[point[0]+p_occ_s*3/20,point[1]+p_occ_s]])
        occupied_points_lines.append([[point[0],point[1]-p_occ_s*1.05],[point[0],point[1]+p_occ_s*1.05]])
        occupied_points_lines.append([[point[0]-p_occ_s/4,point[1]],[point[0]+p_occ_s/4,point[1]]])
        occupied_points_lines.append([[point[0]-p_occ_s*3/20,point[1]+p_occ_s],[point[0]-p_occ_s*3/20,point[1]-p_occ_s]])
        occupied_points_lines.append([[point[0]+p_occ_s*3/20,point[1]+p_occ_s],[point[0]+p_occ_s*3/20,point[1]-p_occ_s]])
    return occupied_points_lines

def get_points_around(my_point,chk_radius,offset):
    ##print('\nChecking point',point[2])
    discovered_points = []
    other_points = get_other_points_list(my_point) #+ occupied_dots
    if offset > chk_radius:
        offset = chk_radius - 1

    for angle in all_angles:
        for radius in range(offset,chk_radius):
            distant_x = my_point[0] + labels_aspect_ratio*radius * math.sin(math.radians(angle))
            distant_y = my_point[1] - radius * math.cos(math.radians(angle))

            for other_point in other_points:
                if abs(distant_x-other_point[0]) <= 1 and abs(distant_y-other_point[1]) <= 1:
                    present = False
                    for member in discovered_points:
                        if member[2] == other_point[2] and member[2] not in ['xaxis','yaxis'] and '_label' not in member[2]:
                            present = True
                    if present == False:
                        discovered_points.append([other_point[0],other_point[1],other_point[2],angle])
    return discovered_points


def get_crowd_coef(point_xy):
    crowd_coef = len(get_points_around(point_xy,3,0))*3
    return crowd_coef

def save_graph(input_graph_name,variant_points_xy,p_colors,labels_xy,debug_lines):
    #fig, ax = pyplot.subplots()
    fig = pyplot.figure(figsize=(16,8))
    ax = fig.add_subplot(111)
    ax.set(xlim=(xmin, xmax), ylim=(ymin, ymax))

    # start X, start Y, height
    orange_area = [50,100,100]
    #orange_area = [50,0,200]
    gradient_size = xmax - orange_area[0]
    # Add Orange Gradient
    gradient_start = colour.Color('#FAF000')
    gradient = list(gradient_start.range_to(colour.Color('#DB0000'),gradient_size))
    gradient_str = []
    for g_color in gradient:
        gradient_str.append(str(g_color).strip('<Color ').strip('>'))
    for n,g_color in enumerate(gradient_str):
        ax.add_patch(patches.Rectangle(xy=(orange_area[0]+n, orange_area[1]) ,width=1, height=orange_area[2], linewidth=0, color=g_color, fill=True, alpha=0.6, zorder=1))

    # Debug occupied
    for line in debug_lines:
        ax.annotate('', (line[0][0], line[0][1]), xytext=(line[1][0],line[1][1]), fontsize=font_size, ha='center', arrowprops=dict(arrowstyle='-',color='blue', linewidth=0.4))

    # Debug check arrows
    #for arrow in debug_arrows:
    #    ax.annotate(str(arrow[4]), (arrow[0], arrow[1]), xytext=(arrow[2],arrow[3]), fontsize=font_size, ha='center', arrowprops=dict(arrowstyle='-',color='grey', linewidth=0.4))


    # Debug check preferred angles

    # Debug check preferred angles

    # Place Arrows
    for l,label in enumerate(labels_xy):
        if abs(label[1] - variant_points_xy[l][1]) < 2:
            ax.annotate(' '*len(variant_points_xy[l][2]), (variant_points_xy[l][0], variant_points_xy[l][1]), xytext=(label[0],label[1]), fontsize=font_size, color='black', alpha=0.9, clip_on=False, ha='center', arrowprops=dict(arrowstyle='<-',color='grey', linewidth=0.8),bbox=dict(pad=0, facecolor="none", edgecolor="none"), zorder=3) #linewidth=0.8 pad=-1.2
        else:
            ax.annotate(' '*len(variant_points_xy[l][2]), (variant_points_xy[l][0], variant_points_xy[l][1]), xytext=(label[0],label[1]), fontsize=font_size, color='black', alpha=0.9, clip_on=False, ha='center', arrowprops=dict(arrowstyle='<-',color='grey', linewidth=0.8),bbox=dict(pad=-2.6, facecolor="none", edgecolor="none"), zorder=3) #linewidth=0.8
    # Place Labels
    for l,label in enumerate(labels_xy):
        if abs(label[1] - variant_points_xy[l][1]) < 2:
            ax.annotate(variant_points_xy[l][2], (variant_points_xy[l][0], variant_points_xy[l][1]), xytext=(label[0],label[1]), fontsize=font_size, color='black', alpha=0.9, clip_on=False, ha='center', arrowprops=dict(arrowstyle='<-',color='grey', linewidth=0),bbox=dict(pad=-1.2, facecolor="none", edgecolor="none"), zorder=4) #linewidth=0.8
        else:
            ax.annotate(variant_points_xy[l][2], (variant_points_xy[l][0], variant_points_xy[l][1]), xytext=(label[0],label[1]), fontsize=font_size, color='black', alpha=0.9, clip_on=False, ha='center', arrowprops=dict(arrowstyle='<-',color='grey', linewidth=0),bbox=dict(pad=-2.6, facecolor="none", edgecolor="none"), zorder=4) #linewidth=0.8



    # Draw Points
    ##ax.scatter([element[0] for element in points_xy], [element[1] for element in points_xy])

    for i,point in enumerate(variant_points_xy):
        ax.scatter(point[0], point[1], color = p_colors[i], linewidth = 0.6, edgecolors = 'black', s=25, zorder=2) # color = p_colors[i]/'grey'

    pyplot.axhline(100,color='black',linewidth=0.8,linestyle='--' ,alpha=1, zorder=1) # '#B6B6B6'
    pyplot.axvline(50,color='black',linewidth=0.8,linestyle='--' ,alpha=1, zorder=1)
    # Plot tick params
    pyplot.locator_params(axis='y',nbins=4)
    pyplot.locator_params(axis='x',nbins=4)
    ax.set_xlabel(graph_labels[1], fontsize=18)
    ax.set_ylabel(graph_labels[2], fontsize=18)
    ax.tick_params(axis="x", labelsize=15)
    ax.tick_params(axis="y", labelsize=15)
    pyplot.axes(ax).set_aspect(aspect = graph_aspect_ratio)
    pyplot.title(graph_labels[0], fontsize = 18) #fontsize = 20/14
    pyplot.savefig(str(input_graph_name)+'.png', dpi = 300)

    return


def check_label_overlap_points_and_axis(lines_to_check,point_xy,probe_point):
    label_overlap_points_and_axis = False
    for occ_line in lines_to_check:
            ##if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[point_xy[0],point_xy[1]],[probe_point[0],probe_point[1]]]):
            ##    label_overlap_points_and_axis = True
        if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+text_height],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+text_height]]):
            label_overlap_points_and_axis = True
        if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
            label_overlap_points_and_axis = True
        if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+text_height]]):
            label_overlap_points_and_axis = True
        if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+text_height],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
            label_overlap_points_and_axis = True
    return label_overlap_points_and_axis

def check_label_line_cross_labels_and_their_lines(lines_to_check,point_xy,probe_point):
    label_line_cross_labels_and_their_lines = False
    for occ_line in lines_to_check:
        if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[point_xy[0],point_xy[1]],[probe_point[0],probe_point[1]]]):
            label_line_cross_labels_and_their_lines = True
    return label_line_cross_labels_and_their_lines

def check_label_overlap_labels_and_their_lines(lines_to_check,point_xy,probe_point, label_overlap_points_and_axis):
    if label_overlap_points_and_axis == False:
        label_overlap_labels_and_their_lines = False
        for occ_line in lines_to_check:
            # if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+text_height/2],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+text_height/2]]):
            #     label_overlap_labels_and_their_lines = True
            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+text_height]]):
                label_overlap_labels_and_their_lines = True
            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+text_height],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
                label_overlap_labels_and_their_lines = True
            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+text_height],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+text_height]]):
                label_overlap_labels_and_their_lines = True
            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
                label_overlap_labels_and_their_lines = True
            # if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+text_height],[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]]]):
            #     label_overlap_labels_and_their_lines = True
            # if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+text_height],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
            #     label_overlap_labels_and_their_lines = True
    else:
        label_overlap_labels_and_their_lines = False
    return label_overlap_labels_and_their_lines

def check_label_line_crosses_other_points(point_xy,probe_point):
    label_line_crosses_other_points = False
    for other_point in get_other_points_list(point_xy):
        if abs(other_point[0] - point_xy[0]) > 1 or abs(other_point[1] - point_xy[1]) > 1:
            occ_lines = occupy_points_lines([other_point])
            for occ_line in occ_lines:
                if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[point_xy[0],point_xy[1]],[probe_point[0],probe_point[1]]]):
                    label_line_crosses_other_points = True
        ##else:
        ##    print('Points',point_xy[2],'and',other_point[2],'are close:',point_xy[0],other_point[0],point_xy[1],other_point[1])
    return label_line_crosses_other_points

# Determine optimal Labels placement locations
# Global Iterations

def check_probe_point_outside(probe_point,point_xy):
    if  probe_point[0] > 0+label_size(point_xy[2])/2 and probe_point[0] < xmax-label_size(point_xy[2])/2 and probe_point[1] > 0+2 and probe_point[1] < ymax-2:
        probe_point_outside = False
    else:
        probe_point_outside = True
    return probe_point_outside

def process_graph_points(i_points_xy,v,simulations_number):
    #simulations_number = 3 #5
    global points_xy
    points_xy = i_points_xy

    #simulation_min_radius = 12
    simulation_radius_step = 4
    simulations_results = []
    ugliness_history = []
    debug_lines = []
    for s in range(0,simulations_number):
        print('\nSimultaion',v+1,s+1,'\n')
        countr = 0
        ##bar_offset = v*simulations_number + s
        ##bar = tqdm(total=len(i_points_xy), position=bar_offset, desc='Simultaion '+str(v+1)+'-'+str(s+1))
        iter_results = []
        iter_occupied_dots = []
        iter_occupied_lines = []
        iter_debug_lines = []
        iter_occupied_conn_lines = []
        iter_labels_xy = []
        ugliness = 0
        for p,point_xy in enumerate(i_points_xy):
            occupied_points_lines = occupy_points_lines(get_other_points_list(point_xy))
            my_point_lines = occupy_points_lines([point_xy])
            print('\nPlacing label for',point_xy[2])
            label_coordinates = []
            label_angle_data_all = determine_label_preferred_angle(point_xy)
            label_angle_data_sorted = sorted(label_angle_data_all, key = lambda x: x[6], reverse = True)
            simulation_min_radius = 12 + get_crowd_coef(point_xy)
            for l,label_angle_data in enumerate(label_angle_data_sorted):
                print('Checking label data',l,label_angle_data[3],label_angle_data[4])
                ##print('Label angle data',label_angle_data[3],label_angle_data[4],'Iter',l)
                ##print('Angle width',label_angle_data[5])
                angles_in_between = get_angles_between(label_angle_data[1],label_angle_data[2])
                #print('angles between',angles_in_between)
                shortest_radius_length = 300
                shortest_radius_angle = 0
                shortest_radius_axis_crossed_coef = 0

                for angle in angles_in_between:
                    angle = normalize_angle(angle)
                    able_to_put = False
                    delta_radius = simulation_min_radius + s*simulation_radius_step # 25

                    while able_to_put != True:

                        probe_point = [point_xy[0] + delta_radius * labels_aspect_ratio*math.sin(math.radians(angle)),point_xy[1] - delta_radius * math.cos(math.radians(angle))]
                        ## Scan Rays Debug
                        ##if point_xy[2] == 'xp7-6mega-0' and l == 0:
                        ##    iter_debug_lines.append([point_xy,probe_point])

                        ##delta_radius_length = get_points_distance(point_xy,probe_point)
                        delta_radius_length = delta_radius

                        label_overlap_points_and_axis = check_label_overlap_points_and_axis(occupied_points_lines + my_point_lines + occupied_axis, point_xy, probe_point)
                        if label_overlap_points_and_axis == False:
                            label_overlap_labels_and_their_lines = check_label_overlap_labels_and_their_lines(iter_occupied_lines+iter_occupied_conn_lines, point_xy, probe_point, label_overlap_points_and_axis)
                            if label_overlap_labels_and_their_lines == False:
                                able_to_put = True
                                probe_point_outside = check_probe_point_outside(probe_point,point_xy)
                                if  probe_point_outside == False:
                                    label_line_crosses_other_points = check_label_line_crosses_other_points(point_xy,probe_point)
                                    if label_line_crosses_other_points == False:
                                        label_line_cross_labels_and_their_lines = check_label_line_cross_labels_and_their_lines(iter_occupied_lines+iter_occupied_conn_lines, point_xy, probe_point)
                                        if label_line_cross_labels_and_their_lines == False:
                                            if delta_radius_length < shortest_radius_length:
                                                shortest_radius_length = delta_radius_length
                                                shortest_radius_angle = angle
                                                label_coordinates = [probe_point[0],probe_point[1],point_xy[2]]
                            else:
                                delta_radius += radius_delta_step
                        else:
                            delta_radius += radius_delta_step

                if label_coordinates != []:
                    print('Selected Angle:',shortest_radius_angle)
                    break

            # Permissive MODE 1 ! - Second Chance
            if label_coordinates == []:
                #label_coordinates = [point_xy[0],point_xy[1]+2,point_xy[2]]
                print('Unable to place Label without crossing obstacles. Trying to allow some crossings')
                for l,label_angle_data in enumerate(label_angle_data_sorted):
                    print('Checking label data',l,label_angle_data[3],label_angle_data[4])
                    ##print('Label angle data',label_angle_data[3],label_angle_data[4],'Iter',l)
                    ##print('Angle width',label_angle_data[5])
                    angles_in_between = get_angles_between(label_angle_data[1],label_angle_data[2])
                    #print('angles between',angles_in_between)
                    shortest_radius_length = 300
                    shortest_radius_angle = 0
                    shortest_radius_axis_crossed_coef = 0

                    for angle in angles_in_between:
                        angle = normalize_angle(angle)
                        able_to_put = False
                        delta_radius = simulation_min_radius + s*simulation_radius_step # 25

                        while able_to_put != True:

                            probe_point = [point_xy[0] + delta_radius * labels_aspect_ratio*math.sin(math.radians(angle)),point_xy[1] - delta_radius * math.cos(math.radians(angle))]
                            ## Scan Rays Debug
                            # if point_xy[2] == 'xp7-20mega-2' and l == 0:
                            #     iter_debug_lines.append([point_xy,probe_point])

                            #delta_radius_length = get_points_distance(point_xy,probe_point)
                            delta_radius_length = delta_radius

                            label_overlap_points_and_axis = check_label_overlap_points_and_axis(occupied_points_lines + my_point_lines + occupied_axis, point_xy, probe_point)
                            if label_overlap_points_and_axis == False:
                                label_overlap_labels_and_their_lines = check_label_overlap_labels_and_their_lines(iter_occupied_lines+iter_occupied_conn_lines, point_xy, probe_point, label_overlap_points_and_axis)
                                if label_overlap_labels_and_their_lines == False:
                                    able_to_put = True
                                    probe_point_outside = check_probe_point_outside(probe_point,point_xy)
                                    if  probe_point_outside == False:
                                        label_line_crosses_other_points = check_label_line_crosses_other_points(point_xy,probe_point)
                                        if label_line_crosses_other_points == False:
                                            label_line_cross_labels_and_their_lines = check_label_line_cross_labels_and_their_lines(iter_occupied_conn_lines, point_xy, probe_point)
                                            if label_line_cross_labels_and_their_lines == False:
                                                if delta_radius_length < shortest_radius_length:
                                                    shortest_radius_length = delta_radius_length
                                                    shortest_radius_angle = angle
                                                    label_coordinates = [probe_point[0],probe_point[1],point_xy[2]]
                                else:
                                    delta_radius += radius_delta_step
                            else:
                                delta_radius += radius_delta_step

                    if label_coordinates != []:
                        print('Selected Angle:',shortest_radius_angle)
                        break

            # Permissive MODE 2 ! - Third Chance
            if label_coordinates == []:
                #label_coordinates = [point_xy[0],point_xy[1]+2,point_xy[2]]
                print('Unable to place Label without crossing obstacles. Trying to allow more crossings')
                for l,label_angle_data in enumerate(label_angle_data_sorted):
                    print('Checking label data',l,label_angle_data[3],label_angle_data[4])
                    ##print('Label angle data',label_angle_data[3],label_angle_data[4],'Iter',l)
                    ##print('Angle width',label_angle_data[5])
                    angles_in_between = get_angles_between(label_angle_data[1],label_angle_data[2])
                    #print('angles between',angles_in_between)
                    shortest_radius_length = 300
                    shortest_radius_angle = 0
                    shortest_radius_axis_crossed_coef = 0

                    for angle in angles_in_between:
                        angle = normalize_angle(angle)
                        able_to_put = False
                        delta_radius = simulation_min_radius + s*simulation_radius_step # 25

                        while able_to_put != True:

                            probe_point = [point_xy[0] + delta_radius * labels_aspect_ratio*math.sin(math.radians(angle)),point_xy[1] - delta_radius * math.cos(math.radians(angle))]
                            ##delta_radius_length = get_points_distance(point_xy,probe_point)
                            delta_radius_length = delta_radius

                            ## Scan Rays Debug
                            # if point_xy[2] == 'xp7-21sk-2' and l == 0:
                            #     iter_debug_lines.append([point_xy,probe_point])

                            label_overlap_points_and_axis = check_label_overlap_points_and_axis(occupied_points_lines + my_point_lines + occupied_axis, point_xy, probe_point)
                            if label_overlap_points_and_axis == False:
                                label_overlap_labels_and_their_lines = check_label_overlap_labels_and_their_lines(iter_occupied_lines+iter_occupied_conn_lines, point_xy, probe_point, label_overlap_points_and_axis)
                                if label_overlap_labels_and_their_lines == False:
                                    able_to_put = True
                                    probe_point_outside = check_probe_point_outside(probe_point,point_xy)
                                    if  probe_point_outside == False:
                                        if delta_radius_length < shortest_radius_length:
                                            shortest_radius_length = delta_radius_length
                                            shortest_radius_angle = angle
                                            label_coordinates = [probe_point[0],probe_point[1],point_xy[2]]
                                else:
                                    delta_radius += radius_delta_step
                            else:
                                delta_radius += radius_delta_step

                    if label_coordinates != []:
                        print('Selected Angle:',shortest_radius_angle)
                        break

            ## Failover
            # if label_coordinates == []:
            #     label_coordinates = [point_xy[0],point_xy[1]+2,point_xy[2]]

            ##place_and_occupy(point_xy,label_coordinates)
            #for char in range(1,int(label_size(label_coordinates[2]))):
                #iter_occupied_dots.append([label_coordinates[0] + char*char_coef - label_size(label_coordinates[2])/2, label_coordinates[1]+2.5,str(label_coordinates[2])+'_label'])
            # Occupy connector line
            iter_occupied_conn_lines.append([point_xy,label_coordinates])
            iter_occupied_conn_lines.append([[label_coordinates[0]-0.4,label_coordinates[1]],[label_coordinates[0]+0.4,label_coordinates[1]]])
            # Occupy label space
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+text_height],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]+text_height]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+text_height],[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]+text_height],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+text_height],[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] + label_size(label_coordinates[2])/4,label_coordinates[1]+text_height],[label_coordinates[0] + label_size(label_coordinates[2])/4,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+text_height],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]+text_height]])
            iter_occupied_lines.append([[label_coordinates[0],label_coordinates[1]],[label_coordinates[0],label_coordinates[1]+text_height]])

            iter_labels_xy.append([label_coordinates[0], label_coordinates[1], shortest_radius_length])

            ##bar.update(1)
            countr += 1
            print('['+str(countr)+'/'+str(len(i_points_xy))+']','#'*int(shortest_radius_length))

            if debug_mode == True:
                points_colors = colour_points(i_points_xy)
                save_graph(graph_name+'_'+str(v+1)+str(s+1)+str(p+1),i_points_xy,points_colors,iter_labels_xy,iter_debug_lines) # [[]] # occupy_points_lines(i_points_xy)


        if debug_mode == False:
            points_colors = colour_points(i_points_xy)
            save_graph(graph_name+'_'+str(v+1)+str(s+1),i_points_xy,points_colors,iter_labels_xy,[]) # iter_occupied_lines+iter_occupied_conn_lines+occupy_points_lines(i_points_xy)+iter_debug_lines

    return




def produce_kv_graph(input_points_xy,iter_num,input_graph_labels,input_graph_name):

    start_time = datetime.datetime.now()

    #cores = multiprocessing.cpu_count()

    global points_xy
    points_xy = input_points_xy

    global graph_labels
    graph_labels = input_graph_labels

    global graph_name
    graph_name = input_graph_name

    #print('Points:',points_xy)


    #print(graph_args_x)
    #print(graph_args_y)

    #ax.scatter(graph_args_x, graph_args_y)

    #pyplot.scatter([1,2,3,4], [5,1,4,2])
    #pyplot.show()

    #pyplot.scatter(graph_args_x, graph_args_y)


    global all_angles
    for i in range(0,361):
        all_angles.append(i)

    #labels_xy = []

    # reserve space on middle axis to not overlap with labels
    ##for x in range(0,int(xmax/4)):
    ##    occupied_axis.append([x*4,100,'xaxis'])
    ##for y in range(0,int(ymax/4)):
    ##    occupied_axis.append([50,y*4,'yaxis'])

    # Calculate sequence variants for points
    points_xy_seq_variants = []

    # Generate variants
    points_xy_seq_variants.append(make_sequence_max_angle())
    ##points_xy_seq_variants.append(make_sequence_brd_dist())



    simulations_results = []
    ugliness_history = []
    debug_lines_history = []

    print('Processing variations\n')
    # with concurrent.futures.ProcessPoolExecutor(max_workers=cores) as executor:
    #     futures_list = []
    #     for v,variant in enumerate(points_xy_seq_variants):
    #         arg = [variant,v,iter_num]
    #         futures_list.append(executor.submit(process_graph_points, *arg))
    #     for future in concurrent.futures.as_completed(futures_list):
    #         if future.result() != None:
    #             simulations_result = future.result()
    for v,variant in enumerate(points_xy_seq_variants):
             arg = [variant,v,iter_num]
             process_graph_points(*arg)

    end_time = datetime.datetime.now()

    print('\n\n\n\n\n\n\n\n\n\n')
    execution_time = end_time - start_time
    print(execution_time)
    ##save_graph('Graph_01')


# Global Variables

points_xy = []
graph_labels = []
graph_name = 'Blank'

xmin = 0
xmax = 100 #100
ymin = 0
ymax = 200

font_size = 6
text_height = 3.8
p_occ_s = 3
graph_aspect_ratio = 0.2
labels_aspect_ratio = 0.25
# 12*10/5*19/10*6/5
char_coef = font_size/10*graph_aspect_ratio*16/8*1.8
points_distance = 11
axis_label_distance = 5.5
point_to_line_min_dist = 1

all_angles = []

occupied_axis = []
occupied_axis = [[[0,100],[100,100]],[[0,102],[100,102]]]

scan_radius = 100 #50

radius_delta_step = 5

debug_mode = False #True

if __name__ == '__main__':

    input_points_xy = []

    input_wb = openpyxl.load_workbook('3par_calc_input.xlsx')
    input_ws = input_wb["Sheet1"]

    for row in input_ws.iter_rows(min_col=2, max_col=27, min_row=3):
        whole_row = []
        for cell in row:
            whole_row.append(cell.value)
        if whole_row[0] != None and whole_row[24] != None and whole_row[25] != None :
            input_points_xy.append([whole_row[24],whole_row[25],whole_row[0]])

    produce_kv_graph(input_points_xy,1)
