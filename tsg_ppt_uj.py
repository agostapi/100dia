#-*- coding: utf-8 -*-

###############TODO##############
#org - short + long version
#5 szabály tesztelés, kiiratni kieso orgokat level3on lesz ilyen
#test-4levelppts

#2. dia - táblázat helyett template
#egység fölötti összes szint eredménye + egy szintig az alatta levő eredmények (org-onként)
#százalékos arányok

#3.dia
#csak az adott organizáció eredménye minden kérdésre
#az előző évihez képest mennyit változott százalékban

#4.dia
#első kérdés minden szervezeti egységre (fölötte összes szint, alatta egy szint)
#előző évhez képest %-os változás a 4-5 válaszra
#1. az 1-2 válasz, 2. a 3. válasz, 3. a 4-5 válasz

#6-7 dia
#marad az 5 válasz
#egység fölötti összes szint, egység alatt egy szint
###############/TODO#############
 

import codecs
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_TICK_MARK
from pptx.enum.chart import XL_LABEL_POSITION
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_LEGEND_POSITION
import time
import structure
import csv
import sys
import os
import shutil

reload(sys)  
sys.setdefaultencoding('utf8')


#UTF8Writer = codecs.getwriter('utf8')

org_short=['TSG','GF OG','VLD','RVL N','VKG 1 N','VKG 7 N','VKG 5 N','VKG 4 N','VKG 3 N','VKG 6 N','VKG 11 N','VKG 2 N','VKG 10 N','VKG 9 N','VKG 8 N','VB Nord','VKG 70 N','SP N','RVL O','VKG 49 O','VKG 50 O','VKG 51 O','VKG 52 O','VKG 53 O','VKG 54 O','VKG 55 O','VKG 56 O','VKG 57 O','VKG 58 O','VKG 74 O','VB Ost','SP O','RVL W','VB West','VKG 14 W','VKG 15 W','VKG 16 W','VKG 17 W','VKG 18 W','VKG 19 W','VKG 20 W','VKG 21 W','VKG 22 W','VKG 23 W','VKG 24 W','VKG 25 W','VKG 71 W','SP W','RVL M','VB Mitte','VKG 61 M','VKG 62 M','VKG 63 M','VKG 64 M','VKG 65 M','VKG 66 M','VKG 67 M','VKG 75 M','SP M','RVL SW','VB Südwest','VKG 28 SW','VKG 29 SW','VKG 30 SW','VKG 31 SW','VKG 32 SW','VKG 33 SW','VKG 34 SW','VKG 35 SW','VKG 36 SW','VKG 72 SW','SP SW','RVL S','VB Süd','VKG 39 S','VKG 40 S','VKG 41 S','VKG 42 S','VKG 43 S','VKG 44 S','VKG 45 S','VKG 46 S','VKG 73 S','VKG 47 S','SP S ','PVD','SO','VSS','HM','RSP','ZM','GF F/CON','F','CORM','HPS','GF HR','HR-DV','HR-MV','HR-BP','BR/SchwbV']

org_long=['Telekom Shop Vertriebsgesellschaft mbH','GF Operatives Geschäft S/M/V','Verkauf Deutschland','Regionale Vertriebsleitung Nord','Verkaufsgebiet 1 Nord Holstein West','Verkaufsgebiet 7 Nord Braunschweig','Verkaufsgebiet 5 Nord Bremen','Verkaufsgebiet 4 Nord Hamburg 2','Verkaufsgebiet 3 Nord Hamburg 1','Verkaufsgebiet 6 Nord Hannover','Verkaufsgebiet 11 Nord Göttingen','Verkaufsgebiet 2 Nord Kiel','Verkaufsgebiet 10 Nord Bielefeld','Verkaufsgebiet 9 Nord Münster','Verkaufsgebiet 8 Nord Osnabrück','Verkaufsbüro Nord','Verkaufsgebiet 70 Nord Partner','Springer Region Nord','Regionale Vertriebsleitung Ost','Verkaufsgebiet 49 Ost Berlin I','Verkaufsgebiet 50 Ost Berlin II','Verkaufsgebiet 51 Ost Berlin III','Verkaufsgebiet 52 Ost Chemnitz','Verkaufsgebiet 53 Ost Dresden','Verkaufsgebiet 54 Ost Erfurt','Verkaufsgebiet 55 Ost Leipzig','Verkaufsgebiet 56 Ost Mecklenburg-Vorpom','Verkaufsgebiet 57 Ost Magdeburg','Verkaufsgebiet 58 Ost Cottbus','Verkaufsgebiet 74 Ost  Partner','Verkaufsbüro Ost','Springer Region Ost','Regionale Vertriebsleitung West','Verkaufsbüro West','Verkaufsgebiet 14 West Wesel','Verkaufsgebiet 15 West Dortmund','Verkaufsgebiet 16 West Soest','Verkaufsgebiet 17 West Duisburg','Verkaufsgebiet 18 West Essen','Verkaufsgebiet 19 West Mönchengladbach','Verkaufsgebiet 20 West Düsseldorf','Verkaufsgebiet 21 West Wuppertal','Verkaufsgebiet 22 West Hagen','Verkaufsgebiet 23 West Aachen','Verkaufsgebiet 24 West Köln','Verkaufsgebiet 25 West Bonn','Verkaufsgebiet 71 West Partner','Springer Region West','Regionale Vertriebsleitung Mitte','Verkaufsbüro Mitte','Verkaufsgebiet 61 Mitte Bad Homburg','Verkaufsgebiet 62 Mitte Darmstadt','Verkaufsgebiet 63 MitteFrankfurt','Verkaufsgebiet 64 MitteGießen','Verkaufsgebiet 65 Mitte Kassel','Verkaufsgebiet 66 Mitte Koblenz','Verkaufsgebiet 67 Mitte Trier','Verkaufsgebiet 75 Mitte Partner','Springer Region Mitte','Regionale Vertriebsleitung SüdWest','Verkaufsbüro Südwest','Verkaufsgebiet 28 Südwest Saarbrücken','Verkaufsgebiet 29 Südwest Mannheim','Verkaufsgebiet 30 Südwest Karlsruhe','Verkaufsgebiet 31 Südwest Heilbronn','Verkaufsgebiet 32 Südwest Stuttgart','Verkaufsgebiet 33 Südwest Freiburg','Verkaufsgebiet 34 Südwest Tübingen','Verkaufsgebiet 35 Südwest Ulm','Verkaufsgebiet 36 Südwest Konstanz','Verkaufsgebiet 72 Südwest Partner','Springer Region Südwest','Regionale Vertriebsleitung Süd','Verkaufsbüro Süd','Verkaufsgebiet 39 Süd Würzburg','Verkaufsgebiet 40 Süd Bayreuth','Verkaufsgebiet 41 Süd Nürnberg','Verkaufsgebiet 42 Süd Regensburg','Verkaufsgebiet 43 Süd Augsburg','Verkaufsgebiet 44 Süd  München','Verkaufsgebiet 45 Süd Kempten','Verkaufsgebiet 46 Süd Rosenheim','Verkaufsgebiet 73 Süd Partner','Verkaufsgebiet 47 Süd Flagship München','Springer Region Süd ','Partner Vertrieb Deutschland','Shop-Oberflächen-Mgt.','Vertriebs-und Servicesteuerung','Handelsmarketing','Retail Strategie & Projekte','Zubehörmanagement','GF Finanzen und Controlling','Finanzen','Compliance & Risk Management','Handelsprozesse und -systeme','GF Personal','HR Development','CC HRM Vertrieb','HR Business Partner','Betriebsrat/Schwerbehindertenvertretung']

#box: default erteket allitunk be, ha negy parametert kap akkor nem lesz box, ha otot akkor hasznalja
#a=the number of the slide
#fill_slide_common(data, tsg_ppt, org_all, a, box=None)
#. az a vminek a vmije first_slidenak azt a shapejet keressuk ami a title (shape es title are fixed by pythonpptx)
def fill_slide_title(tsg_ppt, a, org_all):
  #datum=get_date
  first_slide=tsg_ppt.slides[0]
  common_slide=tsg_ppt.slides[a]
  org_1 = first_slide.placeholders[1]
  org = common_slide.placeholders[17]
  #text_frame = org.text_frame
  #text_frame.clear()
  #p = text_frame.paragraphs[0]
  #run = p.add_run()
  org.text=org_all
  org_1.text = org_all + "\n"+(time.strftime("%d.%m.%Y"))
  #subtitle = first_slide.placeholders[1] 
  #subtitle_2 = common_slide.placeholders[17]
  #subtitle.text=org_all
  return tsg_ppt
  #return -1: if the func doesnwork, we get with this parancs info about the wrong working, egyebkent exit and error code 

def fill_slide_common(tsg_ppt, org_all, a, c1, c2, c3, c4, c5, data, box=None):
#will be fill_slide_common(tsg_ppt, org_all, a):
  common_slide=tsg_ppt.slides[a]
  mittelwert = common_slide.placeholders[18]
  text_frame = mittelwert.text_frame
  text_frame.clear()
  p = text_frame.paragraphs[0]
  run = p.add_run()
  run.text=box
  font = run.font
  font.size=Pt(10.5)
  font.color.rgb = RGBColor(226, 0, 116)

  chart_data = ChartData()
  #chart_data.categories= [c1,c2,c3,c4,c5]
  chart_data.categories=['1a', '2a', 'a3', 'a4']
  #chart_data.add_series('01', (data[0], data[1], data[2], data[3], data[4]))
  a = [0.3,0.4,0.3] 
  b = [0.2,0.3, 0.5]
  c = [0.2,0.2,0.6]
  d = [0.1, 0.1, 0.8]
  #chart_data.add_series('1',(a,b,c))
  chart_data.add_series('2',(a[0], b[0], c[0], d[0]))
  chart_data.add_series('3',(a[1], b[1], c[1], d[1]))
  chart_data.add_series('4',(a[2], b[2], c[2], d[2]))
  x,y,cx,cy = Inches(0.3), Inches(2.25), Inches(9.38), Inches(4.7)
  graphic_frame = common_slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED_100, x, y, cx, cy, chart_data)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  chart.has_legend = True
  chart.plots[0].has_data_labels = True
  data_labels = chart.plots[0].data_labels
  data_labels.number_format = '0%'
  #nehasznaaalddata_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  #value_axis.maximum_scale = 100.0
  tick_labels = value_axis.tick_labels
  tick_labels.number_format = '0%'
  tick_labels.font.size = Pt(12)
  tick_labels.font.type = 'Tele-GroteskEENor'
  chart.legend.position = XL_LEGEND_POSITION.BOTTOM
  chart.legend.font.size = Pt(12)
  chart.legend.font.type = 'Tele-GroteskEENor' 
  #plot = chart.plots[0]
  #plot.has_data_labels = True
  #data_labels = plot.data_labels
  data_labels.font.size = Pt(12)
  #data_labels.number_format = '0"%"'
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  #data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
  category_axis.tick_labels.font.size = Pt(12)
  #bar_plot = chart.plots[0]
  #bar_plot.gap_width = 20
  #bar_plot.overlap = -20
  #print data
  #if (box):
    #print box
  
  chart.replace_data(chart_data)
  return tsg_ppt

def fill_table_slide(tsg_ppt,org_all, n):
  table_slide = tsg_ppt.slides[1]
  #rows = 3
  #cols = 2
  #left = Inches(3.62)
  #top = Inches(3.1)
  #width = Inches(5.0)
  #height = Inches (2.3)
  #table = table_slide.shapes.add_table(rows, cols, left, top, width, height).table
  # set column widths
  #table.columns[0].width = Inches(4.0)
  #table.columns[1].width = Inches(2.0)
  #table.cell(0, 0).text = 'Anzahl der Eingeladenen'
  #table.cell(0, 1).text = str(a)
  #table.cell(1, 0).text = 'Anzahl der Teilnehmer'
  #table.cell(1, 1).text = str(i)
  #table.cell(2, 0).text = 'Quote'
  #table.cell(2, 1).text = n
  chart_data = ChartData()
  chart_data.categories= [org_all]
  chart_data.add_series('01', (n))
  x,y,cx,cy = Inches(0.3), Inches(2.25), Inches(9.38), Inches(4.7)
  graphic_frame = table_slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  value_axis.maximum_scale = 100.0
  tick_labels = value_axis.tick_labels
  tick_labels.number_format = '0"%"'
  tick_labels.font.size = Pt(12)
  tick_labels.font.type = 'Tele-GroteskEENor'
  plot = chart.plots[0]
  plot.has_data_labels = True
  data_labels = plot.data_labels
  data_labels.font.size = Pt(12)
  data_labels.number_format = '0"%"'
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
  category_axis.tick_labels.font.size = Pt(12)
  bar_plot = chart.plots[0]
  bar_plot.gap_width = 20
  bar_plot.overlap = -20
 
  return tsg_ppt


def fill_slide_not_common(tsg_ppt, org_all, a, c1, c2, c3, c4, c5,data ,box=None):
  fill_slide_title(tsg_ppt, a, org_all)
  not_common_slide=tsg_ppt.slides[a]
  mittelwert = not_common_slide.placeholders[18]
  text_frame = mittelwert.text_frame
  text_frame.clear()
  p = text_frame.paragraphs[0]
  run = p.add_run()
  run.text=box
  font = run.font
  font.size=Pt(10.5)
  font.color.rgb = RGBColor(226, 0, 116)

  chart_data = ChartData()
  chart_data.categories= [c1, c2, c3, c4, c5]
  chart_data.add_series('01', (data[0],data[1],data[2],data[3],data[4]))
  x,y,cx,cy = Inches(0.3), Inches(2.25), Inches(9.38), Inches(4.7)
  graphic_frame = not_common_slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  value_axis.maximum_scale = 100.0
  tick_labels = value_axis.tick_labels
  tick_labels.number_format = '0"%"'
  tick_labels.font.size = Pt(12)
  tick_labels.font.type = 'Tele-GroteskEENor'
  plot = chart.plots[0]
  plot.has_data_labels = True
  data_labels = plot.data_labels
  data_labels.font.size = Pt(12)
  data_labels.number_format = '0"%"'
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
  category_axis.tick_labels.font.size = Pt(12)
  bar_plot = chart.plots[0]
  bar_plot.gap_width = 20
  bar_plot.overlap = -20
  #print data
  #if (box):
    #print box
  
  chart.replace_data(chart_data)
  return tsg_ppt

def fill_slide_not_common_2(tsg_ppt, org_all, a, c1, c2, c3, c4, c5,data, box=None):
  fill_slide_title(tsg_ppt, a, org_all)
  not_common_slide=tsg_ppt.slides[a]
  mittelwert = not_common_slide.placeholders[18]
  text_frame = mittelwert.text_frame
  text_frame.clear()
  p = text_frame.paragraphs[0]
  run = p.add_run()
  run.text=box
  font = run.font
  font.size=Pt(10.5)
  font.color.rgb = RGBColor(226, 0, 116)

  chart_data = ChartData()
  chart_data.categories= [c1,c2,c3,c4,c5]
  chart_data.add_series('01', (data[0],data[1],data[2],data[3],data[4]))
  x,y,cx,cy = Inches(0.3), Inches(2.25), Inches(9.38), Inches(4.7)
  graphic_frame = not_common_slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  value_axis.maximum_scale = 100.0
  tick_labels = value_axis.tick_labels
  tick_labels.number_format = '0"%"'
  tick_labels.font.size = Pt(12)
  tick_labels.font.type = 'Tele-GroteskEENor'
  plot = chart.plots[0]
  plot.has_data_labels = True
  data_labels = plot.data_labels
  data_labels.font.size = Pt(12)
  data_labels.number_format = '0"%"'
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
  category_axis.tick_labels.font.size = Pt(12)
  bar_plot = chart.plots[0]
  bar_plot.gap_width = 20
  bar_plot.overlap = -20
  #print data
  #if (box):
    #print box
  
  chart.replace_data(chart_data)
  return tsg_ppt
def fill_slide_mean(tsg_ppt, org_all, a, a1, a2, a3, a4, a5, a6, a7, data, box=None):
  common_slide=tsg_ppt.slides[a]
  textbox=common_slide.shapes[0]
  chart_data = ChartData()
  chart_data.categories= [a1,a2,a3,a4,a5,a6,a7]
  chart_data.add_series('01', (data[8], data[7], data[6], data[5], data[4], data[3], data[0]))
  x,y,cx,cy = Inches(0.3), Inches(2.25), Inches(9.38), Inches(4.7)
  graphic_frame = common_slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  value_axis.maximum_scale = 5
  value_axis.minimum_scale = 1
  tick_labels = value_axis.tick_labels
  tick_labels.number_format = '0.00'
  tick_labels.font.size = Pt(12)
  tick_labels.font.type = 'Tele-GroteskEENor'
  plot = chart.plots[0]
  plot.has_data_labels = True
  data_labels = plot.data_labels
  data_labels.font.size = Pt(12)
  data_labels.number_format = '0.00'
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
  category_axis.tick_labels.font.size = Pt(12)
  bar_plot = chart.plots[0]
  bar_plot.gap_width = 20
  #bar_series = chart.series.bar_series
  #bar_series.fill.solid()
  #bar_series.fill.color.rgb = RGBColor(0, 0, 0)
  

  common_slide=tsg_ppt.slides[a]
  return tsg_ppt

#org is the organization, lista is the list on level org-1, listb is the list with the answers
def getchildresults(org, lista, listb): 
  children = []
  results = []
  for child in structure.orgstructure[org]['child']:
    children.append(child)
  print(children)
  for child in children:
    results.append(listb[lista.index(child)])
    #print(listb.index(child))
  return children, results
    









def id_to_questiontexts(v_id):
  question_ids=['v_175', 'v_176', 'v_177', 'v_179', 'v_180', 'v_181', 'v_182', 'v_183', 'v_184']
  question_texts=["Ich erhalte von meiner Führungskraft regelmäßig Rückmeldungen zu meiner Leistung.", "Wie häufig erhalte ich Rückmeldung zu meiner Leistung von meiner Führungskraft?","Wie lange dauert das Feedbackgespräch durchschnittlich?","Ich erhalte Rückmeldungen zu meiner Gesprächsführung im Kundenkontakt (interner/externer Kunde).", "Ich erhalte Rückmeldungen zu meiner Gesprächsführung im Kundenkontakt (interner/externer Kunde).", "Das Feedback hilft mir, mein Verhalten zu verändern.", "Ich erhalte Feedback zu meinem Beitrag zum Teamerfolg.", "Am Ende des Feedbackgesprächs werden Absprachen getroffen.", "Ich gebe meiner Führungskraft Feedback."]
  for a in range(0,8):
    if (question_ids[a]==v_id):
      return question_texts[a]
  return -1




#fill_slide_common('asdf',1,1,1)

#def asdf(data):
#  mylist = list()
#  for rows in data:
#    if rows[i] not in data and rows[i] and rows[i] not '-99':


##################MAIN###########################

tsg_ppt=Presentation('tsg_templ_3.pptx')

#while (van_graph_element):
#fill_slide_title(tsg_ppt, 2, "Telekom Shop Vertriebsgesellschaft mbH")
##fill_slide_common(data, tsg_ppt, org_all, a, box=None)
mean='2.75'
gultigantwort='600'
#fill_slide_common(1, tsg_ppt, "Telekom Shop Vertriebsgesellschaft mbH", 3, "Trifft überhaupt nicht zu", "Trifft eher nicht zu", "Teils-teils", "trifft eher zu", "Trifft voll zu",  'Mittelwert auf fünfstufiger Skala:'+'\n'+'2,79'+'\n'+'N=300')
org_text_short="TSG"

#--------
if(os.path.isdir('out2')):
  try:
    shutil.rmtree('out2')
  except OSError, e:
    sys.exit("valami fos... in deleting out dir\nError: %s - %s." % (e.filename,e.strerror))
#create out dir, if exists
try:
  os.mkdir('out2')
except OSError, e:
  sys.exit("valami fos... in creating out dir\nError: %s - %s." % (e.filename,e.strerror))



file1_in = open(sys.argv[1], 'rU')
try:
  data_reader = csv.reader((line.replace('\0','') for line in file1_in), delimiter=';')
  my_content = list(data_reader)
  my_header = my_content[0]
  del my_content[0]
finally:
  file1_in.close()

structure.fill_in_org()
print str(len(org_long)) + " " + str(len(org_short))

############
#Fill in org lists

level5_users,level5_numbers,level5_filledin_users = structure.fill_in_users(structure.list5, "u_org_5", structure.users_number5)
list_full_5 = (structure.printout(structure.list5,level5_users))
filled_list_5,my_sums_5,my_means_5 = structure.count_percent(list_full_5, level5_numbers)

level4_users,level4_numbers,level4_filledin_users = structure.fill_in_users(structure.list4, "u_org_4", structure.users_number4)
list_full_4 = (structure.printout(structure.list4,level4_users))
filled_list_4,my_sums_4,my_means_4 = structure.count_percent(list_full_4, level4_numbers)

level3_users,level3_numbers,level3_filledin_users = structure.fill_in_users(structure.list3, "u_org_3", structure.users_number3)
list_full_3 = (structure.printout(structure.list3,level3_users))
filled_list_3,my_sums_3,my_means_3 = structure.count_percent(list_full_3, level3_numbers)

level2_users,level2_numbers,level2_filledin_users = structure.fill_in_users(structure.list2, "u_org_2", structure.users_number2)
list_full_2 = (structure.printout(structure.list2,level2_users))
filled_list_2,my_sums_2,my_means_2 = structure.count_percent(list_full_2, level2_numbers)

level1_users,level1_numbers,level1_filledin_users = structure.fill_in_users(structure.list1, "u_org_1", structure.users_number1)
list_full_1 = (structure.printout(structure.list1,level1_users))
filled_list_1,my_sums_1,my_means_1 = structure.count_percent(list_full_1, level1_numbers)

#print(filled_list_2)
#print(structure.list2)
#lista1, lista2 = getchildresults('TSG', structure.list2, filled_list_2)
#print(structure.list2)
#print(lista1, lista2)

def fill_in_orgstruct_questions():

  for org in structure.list1:
    structure.orgstructure[org]['q1'] = filled_list_1[structure.list1.index(org)][0]
    structure.orgstructure[org]['q2'] = filled_list_1[structure.list1.index(org)][1]
    structure.orgstructure[org]['q3'] = filled_list_1[structure.list1.index(org)][2]
    structure.orgstructure[org]['q4'] = filled_list_1[structure.list1.index(org)][3]
    structure.orgstructure[org]['q5'] = filled_list_1[structure.list1.index(org)][4]
    structure.orgstructure[org]['q6'] = filled_list_1[structure.list1.index(org)][5]
    structure.orgstructure[org]['q7'] = filled_list_1[structure.list1.index(org)][6]
    structure.orgstructure[org]['q8'] = filled_list_1[structure.list1.index(org)][7]
    structure.orgstructure[org]['q9'] = filled_list_1[structure.list1.index(org)][8]

  for org in structure.list2:
    structure.orgstructure[org]['q1'] = filled_list_2[structure.list2.index(org)][0]
    structure.orgstructure[org]['q2'] = filled_list_2[structure.list2.index(org)][1]
    structure.orgstructure[org]['q3'] = filled_list_2[structure.list2.index(org)][2]
    structure.orgstructure[org]['q4'] = filled_list_2[structure.list2.index(org)][3]
    structure.orgstructure[org]['q5'] = filled_list_2[structure.list2.index(org)][4]
    structure.orgstructure[org]['q6'] = filled_list_2[structure.list2.index(org)][5]
    structure.orgstructure[org]['q7'] = filled_list_2[structure.list2.index(org)][6]
    structure.orgstructure[org]['q8'] = filled_list_2[structure.list2.index(org)][7]
    structure.orgstructure[org]['q9'] = filled_list_2[structure.list2.index(org)][8]

  for org in structure.list3:
    structure.orgstructure[org]['q1'] = filled_list_3[structure.list3.index(org)][0]
    structure.orgstructure[org]['q2'] = filled_list_3[structure.list3.index(org)][1]
    structure.orgstructure[org]['q3'] = filled_list_3[structure.list3.index(org)][2]
    structure.orgstructure[org]['q4'] = filled_list_3[structure.list3.index(org)][3]
    structure.orgstructure[org]['q5'] = filled_list_3[structure.list3.index(org)][4]
    structure.orgstructure[org]['q6'] = filled_list_3[structure.list3.index(org)][5]
    structure.orgstructure[org]['q7'] = filled_list_3[structure.list3.index(org)][6]
    structure.orgstructure[org]['q8'] = filled_list_3[structure.list3.index(org)][7]
    structure.orgstructure[org]['q9'] = filled_list_3[structure.list3.index(org)][8]

  for org in structure.list4:
    structure.orgstructure[org]['q1'] = filled_list_4[structure.list4.index(org)][0]
    structure.orgstructure[org]['q2'] = filled_list_4[structure.list4.index(org)][1]
    structure.orgstructure[org]['q3'] = filled_list_4[structure.list4.index(org)][2]
    structure.orgstructure[org]['q4'] = filled_list_4[structure.list4.index(org)][3]
    structure.orgstructure[org]['q5'] = filled_list_4[structure.list4.index(org)][4]
    structure.orgstructure[org]['q6'] = filled_list_4[structure.list4.index(org)][5]
    structure.orgstructure[org]['q7'] = filled_list_4[structure.list4.index(org)][6]
    structure.orgstructure[org]['q8'] = filled_list_4[structure.list4.index(org)][7]
    structure.orgstructure[org]['q9'] = filled_list_4[structure.list4.index(org)][8]

  for org in structure.list5:
    structure.orgstructure[org]['q1'] = filled_list_5[structure.list5.index(org)][0]
    structure.orgstructure[org]['q2'] = filled_list_5[structure.list5.index(org)][1]
    structure.orgstructure[org]['q3'] = filled_list_5[structure.list5.index(org)][2]
    structure.orgstructure[org]['q4'] = filled_list_5[structure.list5.index(org)][3]
    structure.orgstructure[org]['q5'] = filled_list_5[structure.list5.index(org)][4]
    structure.orgstructure[org]['q6'] = filled_list_5[structure.list5.index(org)][5]
    structure.orgstructure[org]['q7'] = filled_list_5[structure.list5.index(org)][6]
    structure.orgstructure[org]['q8'] = filled_list_5[structure.list5.index(org)][7]
    structure.orgstructure[org]['q9'] = filled_list_5[structure.list5.index(org)][8]



fill_in_orgstruct_questions()
#print(structure.orgstructure)

for rows in structure.orgstructure.keys():
  print(rows + ": ", structure.orgstructure[rows])
print(filled_list_2[0][0])
#
######
#create ppt for 5th level
for asdf,my_organi in enumerate(structure.list5):
  if (level5_filledin_users[asdf] < 5):
    print 'not created: '+my_organi
  else:
    tsg_ppt=Presentation('tsg_templ.pptx')
    fill_slide_mean(tsg_ppt, my_organi,  2, "Ich gebe meiner Führungskraft Feedback.", "Am Ende des Feedbackgesprächs werden Absprachen getroffen.", "Ich erhalte Feedback zu meinem Beitrag zum Teamerfolg.", "Das Feedback hilft mir, mein Verhalten zu verändern.", "Das Feedbackgespräch baut auf vorherigem Feedback auf.", "Ich erhalte Rückmeldungen zu meiner Gesprächsführung im Kundenkontakt (interner/externer Kunde).", "Wie häufig erhalte ich Rückmeldung zu meiner Leistung von meiner Führungskraft?", my_means_5[asdf])
    fill_table_slide(tsg_ppt, structure.list5[asdf], str(round(float(level5_filledin_users[asdf] / float(level5_numbers[asdf]))*100,2)) + '%')
    for i in range(1, 12):
      fill_slide_title(tsg_ppt, i, org_long[org_short.index(my_organi)]) #org_long_name
    for i in range(0, 12):
      if (6 <= i <=12 or i==3):
        fill_slide_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "Trifft überhaupt nicht zu", "Trifft eher nicht zu", "Teils-teils", "Trifft eher zu", "Trifft voll zu", filled_list_5[asdf][i-3],'Mittelwert auf fünfstufiger Skala:'+'\n'+str(my_means_5[asdf][i-3])+'\n'+"Gültige Antworten:"+'\n'+str(my_sums_5[asdf][i-3])) #org_long_name
        #will be:
        #1. tsg_ppt
        #2. org name
        #3. number of slide
        #??? pass other orgs, or search in function??
        #fill_slide_common(tsg_ppt, org_long[org_short.index(my_organi)], i)
      elif (i==4):
        fill_slide_not_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "täglich", "maximal 1x pro Woche", "bis zu 1x pro Monat", "halbjährlich", "seltener", filled_list_5[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_5[asdf][i-3]))
      elif (i==5):
        fill_slide_not_common_2(tsg_ppt, org_long[org_short.index(my_organi)], i, "1-3 min", "3-5 min", "5-15 min", "15-30 min", "länger", filled_list_5[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_5[asdf][i-3]))
    tsg_ppt.save("out1/"+(structure.list5[asdf].replace(" ", "_")).replace("/","_")+"_TSG_Leadership_Survey"+".pptx")

######
#create ppt for 4th level
for asdf,my_organi in enumerate(structure.list4):
  if (level4_filledin_users[asdf] < 5):
    print 'not created: '+my_organi
  else:
    tsg_ppt=Presentation('tsg_templ.pptx')
    fill_slide_mean(tsg_ppt, my_organi,  2, "Ich gebe meiner Führungskraft Feedback.", "Am Ende des Feedbackgesprächs werden Absprachen getroffen.", "Ich erhalte Feedback zu meinem Beitrag zum Teamerfolg.", "Das Feedback hilft mir, mein Verhalten zu verändern.", "Das Feedbackgespräch baut auf vorherigem Feedback auf.", "Ich erhalte Rückmeldungen zu meiner Gesprächsführung im Kundenkontakt (interner/externer Kunde).", "Wie häufig erhalte ich Rückmeldung zu meiner Leistung von meiner Führungskraft?",my_means_4[asdf])
    fill_table_slide(tsg_ppt, level4_numbers[asdf], level4_filledin_users[asdf], str(round(float(level4_filledin_users[asdf] / float(level4_numbers[asdf]))*100,2)) + '%')
    for i in range(1, 12):
      fill_slide_title(tsg_ppt, i, org_long[org_short.index(my_organi)]) #org_long_name
    for i in range(0, 12):
      if (6 <= i <=12 or i==3):
        fill_slide_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "Trifft überhaupt nicht zu", "Trifft eher nicht zu", "Teils-teils", "Trifft eher zu", "Trifft voll zu", filled_list_4[asdf][i-3],'Mittelwert auf fünfstufiger Skala:'+'\n'+str(my_means_4[asdf][i-3])+'\n'+"Gültige Antworten:"+'\n'+str(my_sums_4[asdf][i-3])) #org_long_name
      elif (i==4):
        fill_slide_not_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "täglich", "maximal 1x pro Woche", "bis zu 1x pro Monat", "halbjährlich", "seltener", filled_list_4[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_4[asdf][i-3]))
      elif (i==5):
        fill_slide_not_common_2(tsg_ppt, org_long[org_short.index(my_organi)], i, "1-3 min", "3-5 min", "5-15 min", "15-30 min", "länger", filled_list_4[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_4[asdf][i-3]))
    tsg_ppt.save("out1/"+(structure.list4[asdf].replace(" ", "_")).replace("/","_")+"_TSG_Leadership_Survey"+".pptx")

######
#create ppt for 3th level
for asdf,my_organi in enumerate(structure.list3):
  if (level3_filledin_users[asdf] < 5):
    print 'not created: '+my_organi
  else: 
    tsg_ppt=Presentation('tsg_templ.pptx')
    fill_slide_mean(tsg_ppt, my_organi,  2, "Ich gebe meiner Führungskraft Feedback.", "Am Ende des Feedbackgesprächs werden Absprachen getroffen.", "Ich erhalte Feedback zu meinem Beitrag zum Teamerfolg.", "Das Feedback hilft mir, mein Verhalten zu verändern.", "Das Feedbackgespräch baut auf vorherigem Feedback auf.", "Ich erhalte Rückmeldungen zu meiner Gesprächsführung im Kundenkontakt (interner/externer Kunde).", "Wie häufig erhalte ich Rückmeldung zu meiner Leistung von meiner Führungskraft?",my_means_3[asdf])
    fill_table_slide(tsg_ppt, level3_numbers[asdf], level3_filledin_users[asdf], str(round(float(level3_filledin_users[asdf] / float(level3_numbers[asdf]))*100,2)) + '%')
    for i in range(1, 12):
      fill_slide_title(tsg_ppt, i, org_long[org_short.index(my_organi)]) #org_long_name
    for i in range(0, 12):
      if (6 <= i <=12 or i==3):
        fill_slide_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "Trifft überhaupt nicht zu", "Trifft eher nicht zu", "Teils-teils", "Trifft eher zu", "Trifft voll zu", filled_list_3[asdf][i-3],'Mittelwert auf fünfstufiger Skala:'+'\n'+str(my_means_3[asdf][i-3])+'\n'+"Gültige Antworten:"+'\n'+str(my_sums_3[asdf][i-3])) #org_long_name
      elif (i==4):
        fill_slide_not_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "täglich", "maximal 1x pro Woche", "bis zu 1x pro Monat", "halbjährlich", "seltener", filled_list_3[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_3[asdf][i-3]))
      elif (i==5):
        fill_slide_not_common_2(tsg_ppt, org_long[org_short.index(my_organi)], i, "1-3 min", "3-5 min", "5-15 min", "15-30 min", "länger", filled_list_3[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_3[asdf][i-3]))
    tsg_ppt.save("out1/"+(structure.list3[asdf].replace(" ", "_")).replace("/","_")+"_TSG_Leadership_Survey"+".pptx")

######
#create ppt for 2th level
for asdf,my_organi in enumerate(structure.list2):
  if (level2_filledin_users[asdf] < 5):
    print 'not created: '+my_organi
  else:
    tsg_ppt=Presentation('tsg_templ.pptx')
    fill_slide_mean(tsg_ppt, my_organi,  2, "Ich gebe meiner Führungskraft Feedback.", "Am Ende des Feedbackgesprächs werden Absprachen getroffen.", "Ich erhalte Feedback zu meinem Beitrag zum Teamerfolg.", "Das Feedback hilft mir, mein Verhalten zu verändern.", "Das Feedbackgespräch baut auf vorherigem Feedback auf.", "Ich erhalte Rückmeldungen zu meiner Gesprächsführung im Kundenkontakt (interner/externer Kunde).", "Wie häufig erhalte ich Rückmeldung zu meiner Leistung von meiner Führungskraft?",my_means_2[asdf])
    fill_table_slide(tsg_ppt, level2_numbers[asdf], level2_filledin_users[asdf], str(round(float(level2_filledin_users[asdf] / float(level2_numbers[asdf]))*100,2)) + '%')
    for i in range(1, 12):
      fill_slide_title(tsg_ppt, i, org_long[org_short.index(my_organi)]) #org_long_name
    for i in range(0, 12):
      if (6 <= i <=12 or i==3):
        fill_slide_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "Trifft überhaupt nicht zu", "Trifft eher nicht zu", "Teils-teils", "Trifft eher zu", "Trifft voll zu", filled_list_2[asdf][i-3],'Mittelwert auf fünfstufiger Skala:'+'\n'+str(my_means_2[asdf][i-3])+'\n'+"Gültige Antworten:"+'\n'+str(my_sums_2[asdf][i-3])) #org_long_name
      elif (i==4):
        fill_slide_not_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "täglich", "maximal 1x pro Woche", "bis zu 1x pro Monat", "halbjährlich", "seltener", filled_list_2[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_2[asdf][i-3]))
      elif (i==5):
        fill_slide_not_common_2(tsg_ppt, org_long[org_short.index(my_organi)], i, "1-3 min", "3-5 min", "5-15 min", "15-30 min", "länger", filled_list_2[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_2[asdf][i-3]))
    tsg_ppt.save("out1/"+(structure.list2[asdf].replace(" ", "_")).replace("/","_")+"_TSG_Leadership_Survey"+".pptx")

#####
#create ppt for 1th level
for asdf,my_organi in enumerate(structure.list1):
  if (level1_filledin_users[asdf] < 5):
    print 'not created: '+my_organi
  else:
    tsg_ppt=Presentation('tsg_templ.pptx')
    fill_slide_mean(tsg_ppt, my_organi,  2, "Ich gebe meiner Führungskraft Feedback.", "Am Ende des Feedbackgesprächs werden Absprachen getroffen.", "Ich erhalte Feedback zu meinem Beitrag zum Teamerfolg.", "Das Feedback hilft mir, mein Verhalten zu verändern.", "Das Feedbackgespräch baut auf vorherigem Feedback auf.", "Ich erhalte Rückmeldungen zu meiner Gesprächsführung im Kundenkontakt (interner/externer Kunde).", "Wie häufig erhalte ich Rückmeldung zu meiner Leistung von meiner Führungskraft?",my_means_1[asdf])
    fill_table_slide(tsg_ppt, level1_numbers[asdf], level1_filledin_users[asdf], str(round(float(level1_filledin_users[asdf] / float(level1_numbers[asdf]))*100,2)) + '%')
    #fill_slide_not_common(1, tsg_ppt, "Telekom Shop Vertriebgesellschaft mbH", 4, "täglich", "maximal 1x pro Woche", "bis zu 1x pro Monat", "halbjährlich", "seltener", "\n"+"\n"+"N=289")
    for i in range(1, 12):
      fill_slide_title(tsg_ppt, i, org_long[org_short.index(my_organi)]) #org_long_name
    for i in range(0, 12):
      if (6 <= i <=12 or i==3):
        fill_slide_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "Trifft überhaupt nicht zu", "Trifft eher nicht zu", "Teils-teils", "Trifft eher zu", "Trifft voll zu", filled_list_1[asdf][i-3],'Mittelwert auf fünfstufiger Skala:'+'\n'+str(my_means_1[asdf][i-3])+'\n'+"Gültige Antworten:"+'\n'+str(my_sums_1[asdf][i-3])) #org_long_name
      elif (i==4):
        fill_slide_not_common(tsg_ppt, org_long[org_short.index(my_organi)], i, "täglich", "maximal 1x pro Woche", "bis zu 1x pro Monat", "halbjährlich", "seltener", filled_list_1[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_1[asdf][i-3]))
      elif (i==5):
        fill_slide_not_common_2(tsg_ppt, org_long[org_short.index(my_organi)], i, "1-3 min", "3-5 min", "5-15 min", "15-30 min", "länger", filled_list_1[asdf][i-3],"\n"+"Gültige Antworten:"+"\n"+str(my_sums_1[asdf][i-3]))
    tsg_ppt.save("out1/"+(structure.list1[asdf].replace(" ", "_")).replace("/","_")+"_TSG_Leadership_Survey"+".pptx")

#################MAIN ENDS HERE###################
#dict to fill in from old file in structure.py
#dict_1 = {'ORG1': 73, 'ORG2': 54}
#dict_2 = {'ORG3': 34, 'ORG34': 234}
