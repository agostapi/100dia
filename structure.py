from __future__ import print_function
from operator import itemgetter
import time
import csv
import sys
import re
import operator
import shutil
import os
#from data import *
import itertools
import copy
import orgs


orgstructure = {}

def myprint(d):
  for k, v in d.iteritems():
    if isinstance(v, dict):
      myprint(v)
    else:
      print("{0} : {1}".format(k, v))

#print structure.keys()
#print '\n'
#myprint(structure)
#print traverse_nested_dict(structure)
#print structure[structure.get('list')[0]].get('list')[0]
list1 = list()
list2 = list()
list3 = list()
list4 = list()
list5 = list()
my_header = None
my_content = None
users_number5 = []
users_number4 = []
users_number3 = []
users_number2 = []
users_number1 = []

def fill_in_users(lista, u_org, listaaa): #list5, 'u_org_5'
  my_full_level_users = []
  my_filled_in_users = []
  for i,org in enumerate(lista):
    my_org_users = []
    filled_in_users = []
    all_users_number = 0
    my_filled_in_users_number = 0
    for rows in my_content:
      if rows[my_header.index(u_org)] == org:
        all_users_number += 1
        if rows[my_header.index('dispcode')] == '31' or rows[my_header.index('dispcode')] =='32':
          my_filled_in_users_number += 1
          my_org_users.append([rows[my_header.index('u_email')],rows[my_header.index('v_175')],rows[my_header.index('v_176')],rows[my_header.index('v_177')],rows[my_header.index('v_179')],rows[my_header.index('v_180')],rows[my_header.index('v_181')],rows[my_header.index('v_182')],rows[my_header.index('v_183')],rows[my_header.index('v_184')]])
        #if org == 'VKG 30 SW':
          #print(len(my_org_users))
          #print org
    my_full_level_users.append(my_org_users)
    my_filled_in_users.append(my_filled_in_users_number)
    listaaa.append(all_users_number)
  return my_full_level_users, listaaa, my_filled_in_users
      
        

def fill_in_org():
  for rows in my_content:
    if rows[my_header.index('u_org_1')] not in list1 and rows[my_header.index('u_org_1')] and rows[my_header.index('u_org_1')] != '-99' and orgs.org_last_yr.has_key(rows[my_header.index('u_org_1')]):
      list1.append(rows[my_header.index('u_org_1')])
      #if the org does not exist in structure, add it
      if rows[my_header.index('u_org_1')] not in orgstructure.keys():
        orgstructure.update({rows[my_header.index('u_org_1')] : { 'parent' : None, 'child' : [], 'level' : 1, 'long name' : orgs.org_short_to_long[rows[my_header.index('u_org_1')]]}})
    if rows[my_header.index('u_org_2')] not in list2 and rows[my_header.index('u_org_2')] and rows[my_header.index('u_org_2')] != '-99' and orgs.org_last_yr.has_key(rows[my_header.index('u_org_2')]):
      list2.append(rows[my_header.index('u_org_2')])
      #if the org does not exist in structure, add it with all the parents, and add itself to the parent's child list
      if rows[my_header.index('u_org_2')] not in orgstructure.keys():
        orgstructure.update({rows[my_header.index('u_org_2')] : { 'parent' : rows[my_header.index('u_org_1')], 'child' : [], 'level' : 2, 'long name' : orgs.org_short_to_long[rows[my_header.index('u_org_2')]]}})
        orgstructure[rows[my_header.index('u_org_1')]]['child'].append(rows[my_header.index('u_org_2')])
    if rows[my_header.index('u_org_3')] not in list3 and rows[my_header.index('u_org_3')] and rows[my_header.index('u_org_3')] != '-99' and orgs.org_last_yr.has_key(rows[my_header.index('u_org_3')]):
      list3.append(rows[my_header.index('u_org_3')])
      if rows[my_header.index('u_org_3')] not in orgstructure.keys():
        orgstructure.update({rows[my_header.index('u_org_3')] : { 'parent' : [rows[my_header.index('u_org_2')],rows[my_header.index('u_org_1')]] , 'child' : [], 'level' : 3, 'long name' : orgs.org_short_to_long[rows[my_header.index('u_org_3')]]}})
        orgstructure[rows[my_header.index('u_org_2')]]['child'].append(rows[my_header.index('u_org_3')])
    if rows[my_header.index('u_org_4')] not in list4 and rows[my_header.index('u_org_4')] and rows[my_header.index('u_org_4')] != '-99' and orgs.org_last_yr.has_key(rows[my_header.index('u_org_4')]):
      list4.append(rows[my_header.index('u_org_4')])
      if rows[my_header.index('u_org_4')] not in orgstructure.keys():
        orgstructure.update({rows[my_header.index('u_org_4')] : { 'parent' : [rows[my_header.index('u_org_3')], rows[my_header.index('u_org_2')],rows[my_header.index('u_org_1')]], 'child' : [], 'level' : 4, 'long name' : orgs.org_short_to_long[rows[my_header.index('u_org_4')]]}})
        orgstructure[rows[my_header.index('u_org_3')]]['child'].append(rows[my_header.index('u_org_4')])
    if rows[my_header.index('u_org_5')] not in list5 and rows[my_header.index('u_org_5')] and rows[my_header.index('u_org_5')] != '-99' and orgs.org_last_yr.has_key(rows[my_header.index('u_org_5')]):
      list5.append(rows[my_header.index('u_org_5')])
      if rows[my_header.index('u_org_5')] not in orgstructure.keys():
        orgstructure.update({rows[my_header.index('u_org_5')] : { 'parent' : [rows[my_header.index('u_org_4')], rows[my_header.index('u_org_3')], rows[my_header.index('u_org_2')],rows[my_header.index('u_org_1')]], 'child' : None , 'level' : 5, 'long name' : orgs.org_short_to_long[rows[my_header.index('u_org_5')]]}})
        orgstructure[rows[my_header.index('u_org_4')]]['child'].append(rows[my_header.index('u_org_5')])

  #for rows in orgstructure.keys():
  #print(rows + ": ", orgstructure[rows])

def printout(a,users_list):
  answers_out = [] #[[0,0,0,0,0], [0,0,0,0,0], [0,0,0,0,0], [0,0,0,0,0], [0,0,0,0,0], [0,0,0,0,0], [0,0,0,0,0], [0,0,0,0,0], [0,0,0,0,0]] #0,0,0,0,0]
  #my_temp_asdf = [0,0,0,0,0]
  #for b in users_list:
  #  answers_out.append([:])
  #print(answers)
  for i,user in enumerate(users_list):
    answers = []
    for j in range(0,9):
      answers.append([0,0,0,0,0])
    #for i in range(0,9):
    #  answers_out.append(my_temp_asdf[:])
    #print(answers_out)
    #print(user)
    for k in range(0,9):
      for j in range(0,5):
        answers[k][j] = 0
    for j,user_a in enumerate(user):
      #print(user_a)
      for n in range(1,10):
        #print(user_a[n])
        #if users_list[i][a][n] == '1':
        if user_a[n] == '1':
          answers[n-1][0] += 1
          #answers_out[i][n-1][0] += 1
        elif user_a[n] == '2': #users_list[i][a][n] == '2':
          answers[n-1][1] += 1
          #answers_out[i][n-1][1] += 1
        elif user_a[n] == '3': #users_list[i][a][n] == '3':
          answers[n-1][2] += 1
          #answers_out[i][n-1][2] += 1
        elif user_a[n] == '4': #users_list[i][a][n] == '4':
          answers[n-1][3] += 1
          #answers_out[i][n-1][3] += 1
        elif user_a[n] == '5': #users_list[i][a][n] == '5':
          answers[n-1][4] += 1
          #answers_out[i][n-1][4] += 1
      #if i == 3:
        #print(answers)
    #print(answers)
    #print(answers_out)
    answers_out.append(answers[:])
    #print(answers_out)
    #print(users_list[0][0][1])
    #print(answers)

  #def fill_answers(org):
    
    

  #for a,rows in enumerate(users_list):
  #  answers[a] += rows
    
  #for i in range(0,8):
  #  for j in range(0,8):
  #    answers[i][j] = 0
  #answers = [[] for x in xrange(8)]
  #print(answers)
  #for users in users_list:

  return answers_out
  
def count_percent(list_a, numbers):
  my_sums = []
  my_means = []
  list_b = copy.deepcopy(list_a)
  list_c = copy.deepcopy(list_a)
  for asdf, rows in enumerate(list_a):
    my_temp_sums = []
    my_temp_means = []
    for asdg, questions in enumerate(rows):
      my_sum = 0
      my_mean = 0
      for i in range(0,5):
        my_sum += questions[i]
      for i in range(0,5):
        #list_b[rows][questions][i] = i*list_b[rows][questions][i] +    / my_sum
        if my_sum == 0:
          questions[i] = 0
        else:
          questions[i] = round(( float(questions[i]) / float(my_sum) ) *100,1)
      if my_sum == 0:
        my_mean = 0
      else:
        my_mean = round(float( 1 * list_b[asdf][asdg][0] + 2 * list_b[asdf][asdg][1] + 3 * list_b[asdf][asdg][2] + 4 * list_b[asdf][asdg][3] + 5 * list_b[asdf][asdg][4] ) / float(my_sum),2)
      my_temp_sums.append(my_sum)
      my_temp_means.append(my_mean)
    my_sums.append(my_temp_sums)
    my_means.append(my_temp_means)
    #print(list_b[0][0])
     
      
  return list_a, my_sums, my_means

def sort_orgchild():
  for org in orgstructure.keys():
    if orgstructure[org]['child']:
      orgstructure[org]['child'].sort(reverse=True)

########MAIN################



file1_in = open(sys.argv[1], 'rU')
  
try:
  data_reader = csv.reader((line.replace('\0','') for line in file1_in), delimiter=';')
  my_content = list(data_reader)
  my_header = my_content[0]
  del my_content[0]
finally:
  file1_in.close()
  
#print(my_header)
fill_in_org()
sort_orgchild()
print(orgstructure)
  
#level5_users,level5_numbers = fill_in_users(list5, "u_org_5", users_number5)
#level4_users = fill_in_users(list4, 'u_org_4')
#print(level5_users,level5_numbers)
#print(level4_users[0][0])
#for i,data in enumerate(list5):
#  printout(data,level5_users[i])
#list_full_5 = (printout(list5,level5_users))
#filled_list_5,my_sums_5,my_means_5 = count_percent(list_full_5, level5_numbers)
#print(my_means_5)
#print(filled_list_5)
#print(my_means_5)
#print(filled_list_5)
#print(filled_list_5[3])


#print(level5_users)



#print(users_number5)

#print(level5_users[0][0])



#print('\n',list1,'\n',list2,'\n',list3,'\n',list4,'\n',list5)
#print(len(list1)+len(list2)+len(list3)+len(list4))
