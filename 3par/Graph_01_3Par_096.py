import openpyxl
from matplotlib import pyplot
import math
import sys
from shapely.geometry import LineString
import colour
import numpy as np
import concurrent.futures
import multiprocessing
import tqdm
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
        angles_data_sorted = sorted(angles_data, key = lambda x: x[5], reverse = True)
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
    if max < min:
        max += 360
    number_of_angles = int(((max-angle_offset*2) - min)/detail_level)+1

    middle_angle = (min + max)/2
    angles_centrified.append(middle_angle)
    for iter in range(1,int(number_of_angles/2)+1):
        angles_centrified.append(middle_angle+iter*detail_level)
        angles_centrified.append(middle_angle-iter*detail_level)

    if len(angles_centrified) == 0:
        angles_centrified.append(middle_angle)
    return angles_centrified

def check_lines_cross(line_a_coords,line_b_coords):
    line_a = LineString(line_a_coords)
    line_b = LineString(line_b_coords)
    return line_a.intersects(line_b)

def flatten_coef_decision(angle):
    angle_verthor_coef = abs(math.cos(math.radians(angle)))*abs(math.cos(math.radians(angle))) *abs(math.cos(math.radians(angle)))
    if angle_verthor_coef < 0.05:
        angle_verthor_coef = 0.05
    #print('calculating angle coef',angle,angle_verthor_coef)
    return angle_verthor_coef

def flatten_coef_radius(angle):
    angle_verthor_coef = abs(math.cos(math.radians(angle))) # *abs(math.cos(math.radians(angle)))
    if angle_verthor_coef < 0.2:
        angle_verthor_coef = 0.2
    #print('calculating angle coef',angle,angle_verthor_coef)
    return angle_verthor_coef

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
    tmp_discovered_points = []
    other_points = get_other_points_list(point) #+ occupied_dots
    label_angle_data = []

    for angle in all_angles:
        for radius in range(0,scan_radius):
            distant_x = point[0] + radius * math.sin(math.radians(angle))
            distant_y = point[1] - radius * math.cos(math.radians(angle))

            for other_point in other_points:
                if abs(distant_x-other_point[0]) <= 1 and abs(distant_y-other_point[1]) <= 1:
                    present = False
                    for member in tmp_discovered_points:
                        if member[2] == other_point[2] and member[2] not in ['xaxis','yaxis'] and '_label' not in member[2]:
                            present = True
                    if present == False:
                        tmp_discovered_points.append([other_point[0],other_point[1],other_point[2],angle])

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
            label_angle_data.append([normalize_angle(temp_angle), temp_angle-angle_width/2, temp_angle+angle_width/2, discovered_points[m][2], discovered_points[n][2],angle_width])

        else:
            n = m+1
            ##print('current and next',discovered_points[m][2],discovered_points[m][3],discovered_points[n][2],discovered_points[n][3])
            temp_angle = discovered_points[m][3] + (get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]]))/2

            #Debug
            #print('preprocessed angle width =',discovered_points[n][3] - discovered_points[m][3])
            ##print('precise angle width =',get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]]))
            #Debug
            angle_width = get_angle_precise([discovered_points[m][0],discovered_points[m][1]],[point[0],point[1]],[discovered_points[n][0],discovered_points[n][1]])
            label_angle_data.append([normalize_angle(temp_angle), temp_angle-angle_width/2, temp_angle+angle_width/2, discovered_points[m][2], discovered_points[n][2],angle_width])

        ##print('label angle',label_angle_data)
    # label_angle: recommended angle, angle minimum, angle maximum, first point of angle, second point of angle
    ##print('Angle for label placement:',label_angle_data)
    if len(discovered_points) < 2:
        label_angle_data.append([180, 0, 360, '0', '360', 360])

    return label_angle_data


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
        occupied_points_lines.append([[point[0]-0.6,point[1]+4],[point[0]+0.6,point[1]-4]])
        occupied_points_lines.append([[point[0]-0.6,point[1]-4],[point[0]+0.6,point[1]+4]])
        occupied_points_lines.append([[point[0],point[1]-4.2],[point[0],point[1]+4.2]])
        occupied_points_lines.append([[point[0]-1,point[1]],[point[0]+1,point[1]]])
        occupied_points_lines.append([[point[0]-0.6,point[1]+4],[point[0]-0.6,point[1]-4]])
        occupied_points_lines.append([[point[0]+0.6,point[1]+4],[point[0]+0.6,point[1]-4]])
    return occupied_points_lines


def save_graph(graph_name,variant_points_xy,p_colors,labels_xy,debug_lines):
    #fig, ax = pyplot.subplots()
    fig = pyplot.figure(figsize=(16,8))
    ax = fig.add_subplot(111)
    ax.set(xlim=(xmin, xmax), ylim=(ymin, ymax))

    # Debug occupied
    ##for line in debug_lines:
    ##    ax.annotate('', (line[0][0], line[0][1]), xytext=(line[1][0],line[1][1]), fontsize=font_size, ha='center', arrowprops=dict(arrowstyle='-',color='grey', linewidth=0.4))

    # Debug check arrows
    ##for arrow in debug_arrows:
    ##    ax.annotate(str(arrow[4]), (arrow[0], arrow[1]), xytext=(arrow[2],arrow[3]), fontsize=font_size, ha='center', arrowprops=dict(arrowstyle='-',color='grey', linewidth=0.4))



    # Place Labels
    point_label = 0
    for label in labels_xy:
        ax.annotate(variant_points_xy[point_label][2], (variant_points_xy[point_label][0], variant_points_xy[point_label][1]), xytext=(label[0],label[1]), fontsize=font_size, alpha=0.7, ha='center', arrowprops=dict(arrowstyle='-',color='grey', linewidth=0.8), zorder=3)
        point_label += 1

    # Draw Points
    ##ax.scatter([element[0] for element in points_xy], [element[1] for element in points_xy])

    for i,point in enumerate(variant_points_xy):
        ax.scatter(point[0], point[1], color = p_colors[i], linewidth = 0.6, edgecolors = 'black', s=100, zorder=2)

    pyplot.axhline(100,color='black',linewidth=0.8,linestyle='--' ,alpha=1, zorder=1) # '#B6B6B6'
    pyplot.axvline(50,color='black',linewidth=0.8,linestyle='--' ,alpha=1, zorder=1)
    # Plot tick params
    pyplot.locator_params(axis='y',nbins=4)
    pyplot.locator_params(axis='x',nbins=4)
    ax.set_xlabel('Используемые ресурсы', fontsize=18)
    ax.set_ylabel('Потребность', fontsize=18)
    ax.tick_params(axis="x", labelsize=15)
    ax.tick_params(axis="y", labelsize=15)
    pyplot.axes(ax).set_aspect(aspect = graph_aspect_ratio)
    pyplot.title("OVERPROV VS UTILIZATION (OVERALL)", fontsize = 20)
    pyplot.savefig(str(graph_name)+'.png', dpi = 300)




# Determine optimal Labels placement locations
# Global Iterations

def process_graph_points(i_points_xy,v,simulations_number):
    #simulations_number = 3 #5
    simulation_min_radius = 12
    simulation_radius_step = 4
    simulations_results = []
    ugliness_history = []
    debug_lines = []
    for s in range(0,simulations_number):
        #print('\nSimultaion',v,s,'\n')
        bar_offset = v*simulations_number + s
        bar = tqdm.tqdm(total=len(i_points_xy), position=bar_offset, desc='Simultaion '+str(v+1)+'-'+str(s+1))
        iter_results = []
        iter_occupied_dots = []
        iter_occupied_lines = []
        iter_labels_xy = []
        ugliness = 0
        for point_xy in i_points_xy:
            occupied_points_lines = occupy_points_lines(get_other_points_list(point_xy))
            my_point_lines = occupy_points_lines([point_xy])
            ##print('\nPlacing label for',point_xy[2])
            label_coordinates = []
            label_angle_data_all = determine_label_preferred_angle(point_xy)
            label_angle_data_sorted = sorted(label_angle_data_all, key = lambda x: x[5], reverse = True)
            for l,label_angle_data in enumerate(label_angle_data_sorted):
                ##print('Label angle data',label_angle_data[3],label_angle_data[4],'Iter',l)
                ##print('Angle width',label_angle_data[5])
                angles_in_between = get_angles_between(label_angle_data[1],label_angle_data[2])
                #print('angles between',angles_in_between)
                shortest_radius = 200
                shortest_radius_angle = 0
                shortest_radius_axis_crossed_coef = 0


                for angle in angles_in_between:
                    angle = normalize_angle(angle)
                    ##print('checking angle',angle)
                    able_to_put = False
                    #label_x = 0
                    #label_y = 0
                    delta_radius = (simulation_min_radius + s*simulation_radius_step) * flatten_coef_radius(angle) # 25
                    placing_attempt = 1


                    while able_to_put != True:
                        placing_attempt += 1

                        probe_point = [point_xy[0] + delta_radius * math.sin(math.radians(angle)),point_xy[1] - delta_radius * math.cos(math.radians(angle))]



                        is_occupied_dot = False
                        for occ_line in occupied_points_lines + my_point_lines + occupied_axis:
                                ##if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[point_xy[0],point_xy[1]],[probe_point[0],probe_point[1]]]):
                                ##    is_occupied_dot = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+2],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+2]]):
                                is_occupied_dot = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+4]]):
                                is_occupied_dot = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+4],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
                                is_occupied_dot = True



                        #for occ_dot in iter_occupied_dots + points_xy:
                        #    if abs(probe_point[0] - occ_dot[0]) < points_distance/2 and abs(probe_point[1]  - occ_dot[1]) < points_distance:
                        #        is_occupied_dot = True


                        label_line_crosses_obstacles = False
                        for occ_line in iter_occupied_lines:
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[point_xy[0],point_xy[1]],[probe_point[0],probe_point[1]]]):
                                label_line_crosses_obstacles = True

                        is_occupied_line = False
                        for occ_line in iter_occupied_lines:
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+2],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+2]]):
                                is_occupied_line = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+6]]):
                                is_occupied_line = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+6],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
                                is_occupied_line = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+6],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+6]]):
                                is_occupied_line = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
                                is_occupied_line = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]+6],[probe_point[0]-label_size(point_xy[2])/2,probe_point[1]]]):
                                is_occupied_line = True
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]+6],[probe_point[0]+label_size(point_xy[2])/2,probe_point[1]]]):
                                is_occupied_line = True

                        label_line_crosses_other_points = False
                        #for point in get_other_points_list(point_xy):
                        #    if point_line_dist(point,[[point_xy[0],point_xy[1]],[probe_point[0],probe_point[1]]]) < point_to_line_min_dist:
                        for occ_line in occupied_points_lines:
                            if check_lines_cross([[occ_line[0][0],occ_line[0][1]],[occ_line[1][0],occ_line[1][1]]],[[point_xy[0],point_xy[1]],[probe_point[0],probe_point[1]]]):
                                label_line_crosses_other_points = True


                        if is_occupied_dot == False and is_occupied_line == False :
                            able_to_put = True
                            if  probe_point[0] > 0+label_size(point_xy[2])/2 and probe_point[0] < xmax-label_size(point_xy[2])/2 and probe_point[1] > 0+2 and probe_point[1] < ymax-2 and label_line_crosses_other_points == False and label_line_crosses_obstacles == False: # and delta_radius < 20:
                                ##print('delta radius',delta_radius)
                                # Debug allowed rays
                                ##debug_arrows.append([point_xy[0],point_xy[1],probe_point[0],probe_point[1],''])
                                if delta_radius/flatten_coef_decision(angle) < shortest_radius/flatten_coef_decision(shortest_radius_angle):
                                    shortest_radius = delta_radius
                                    shortest_radius_angle = angle
                                    label_coordinates = [probe_point[0],probe_point[1],point_xy[2]]
                                    ##print('coefficients applied',delta_radius/flatten_coef_decision(angle),'shortest radius',shortest_radius)

                        else:

                            delta_radius += 0.5

                if label_coordinates != []:
                    break

                if l == (len(label_angle_data_all) -1) and label_coordinates == []:
                    label_coordinates = [point_xy[0],point_xy[1]+2,point_xy[2]]
                    ##print('Failing back to initial location..')

            ##place_and_occupy(point_xy,label_coordinates)
            #for char in range(1,int(label_size(label_coordinates[2]))):
                #iter_occupied_dots.append([label_coordinates[0] + char*char_coef - label_size(label_coordinates[2])/2, label_coordinates[1]+2.5,str(label_coordinates[2])+'_label'])
            # Occupy connector line
            iter_occupied_lines.append([point_xy,label_coordinates])
            # Occupy label space
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+4],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]+4]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+4],[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]+4],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+4],[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] + label_size(label_coordinates[2])/4,label_coordinates[1]+4],[label_coordinates[0] + label_size(label_coordinates[2])/4,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+2],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]+2]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]+4],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]]])
            iter_occupied_lines.append([[label_coordinates[0] - label_size(label_coordinates[2])/2,label_coordinates[1]],[label_coordinates[0] + label_size(label_coordinates[2])/2,label_coordinates[1]+4]])
            iter_occupied_lines.append([[label_coordinates[0],label_coordinates[1]],[label_coordinates[0],label_coordinates[1]+4]])

            iter_labels_xy.append([label_coordinates[0], label_coordinates[1], shortest_radius])

            ##print('Placed',shortest_radius)
            bar.update(1)



        #print('Found places:',iter_labels_xy)
        for label_vector in iter_labels_xy:
            if label_vector[2] > 30:
                ugliness += label_vector[2]*10
            elif label_vector[2] > 20:
                ugliness += label_vector[2]*5
            else:
                ugliness += label_vector[2]

        points_colors = colour_points(i_points_xy)

        save_graph('Graph_01_'+str(v+1)+str(s+1),i_points_xy,points_colors,iter_labels_xy,iter_occupied_lines+occupy_points_lines(points_xy))

    return




def produce_graph_01_variants(input_points_xy,iter_num):

    start_time = datetime.datetime.now()

    cores = multiprocessing.cpu_count()

    global points_xy
    points_xy = input_points_xy


    #print('Points:',points_xy)


    #print(graph_args_x)
    #print(graph_args_y)

    global xmin
    global xmax
    global ymin
    global ymax

    global font_size
    global graph_aspect_ratio
    global char_coef
    global points_distance
    global axis_label_distance
    global point_to_line_min_dist

    #ax.scatter(graph_args_x, graph_args_y)

    #pyplot.scatter([1,2,3,4], [5,1,4,2])
    #pyplot.show()

    #pyplot.scatter(graph_args_x, graph_args_y)


    global all_angles
    for i in range(0,361):
        all_angles.append(i)

    #labels_xy = []
    global occupied_axis

    # reserve space on middle axis to not overlap with labels
    ##for x in range(0,int(xmax/4)):
    ##    occupied_axis.append([x*4,100,'xaxis'])
    ##for y in range(0,int(ymax/4)):
    ##    occupied_axis.append([50,y*4,'yaxis'])

    # Calculate sequence variants for points
    points_xy_seq_variants = []

    # Generate variants
    points_xy_seq_variants.append(make_sequence_brd_dist())
    points_xy_seq_variants.append(make_sequence_max_angle())


    simulations_results = []
    ugliness_history = []
    debug_lines_history = []

    print('Processing variations\n')
    with concurrent.futures.ProcessPoolExecutor(max_workers=cores) as executor:
        futures_list = []
        for v,variant in enumerate(points_xy_seq_variants):
            arg = [variant,v,iter_num]
            futures_list.append(executor.submit(process_graph_points, *arg))
        for future in concurrent.futures.as_completed(futures_list):
            if future.result() != None:
                simulations_result = future.result()

    end_time = datetime.datetime.now()

    print('\n\n\n\n\n\n\n\n\n\n')
    execution_time = end_time - start_time
    print(execution_time)
    ##save_graph('Graph_01')


# Global Variables

points_xy = []

xmin = 0
xmax = 100 #100
ymin = 0
ymax = 200

font_size = 8
graph_aspect_ratio = 0.2
# 12*10/5*19/10*6/5
char_coef = font_size/10*graph_aspect_ratio*16/8*1.8
points_distance = 11
axis_label_distance = 5.5
point_to_line_min_dist = 1

all_angles = []

occupied_axis = []
occupied_axis = [[[0,100],[100,100]]]

scan_radius = 50

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

    produce_graph_01_variants(input_points_xy,3)
