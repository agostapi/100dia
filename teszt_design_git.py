#-*- coding: utf-8 -*-

###############TODO##############
#org - short + long version
#5 szabály tesztelés, kiiratni kieso orgokat level3on lesz ilyen
#test-4levelppts
###############/TODO#############
 
from __future__ import absolute_import, print_function, unicode_literals
import codecs
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_TICK_MARK
from pptx.enum.chart import XL_LABEL_POSITION
from pptx.enum.chart import XL_TICK_LABEL_POSITION
from pptx.util import Cm
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.line import LineFormat
from pptx.dml.chtfmt import ChartFormat
from pptx.dml.fill import FillFormat
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_LEGEND_POSITION
import time
#import structure
import csv
import sys
import os
import shutil
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
reload(sys)  
sys.setdefaultencoding('utf8')


#UTF8Writer = codecs.getwriter('utf8')


#fill the titles in all slide withe the long name of the orgs, add datum to the "0" slide

def fill_slide_0_and_titles(tsg_ppt, orglongname):
  for i in range(0,12):
    if i==0:
      first_slide = tsg_ppt.slides[0]
      org_1 = first_slide.placeholders[1]
      org_1.text = orglongname + "\n"+(time.strftime("%d.%m.%Y"))
    else:
      other_slide = tsg_ppt.slides[i]
      org = other_slide.placeholders[17]
      org.text = orglongname
  return tsg_ppt


# fill the first slide with the percents of the participants pro org

def fill_slide_1(tsg_ppt, org_names, org_percents):
  slide = tsg_ppt.slides[1]
  chart_data = ChartData()
  chart_data.categories= org_names
  series=chart_data.add_series('01', org_percents)
  x,y,cx,cy = Inches(0.3), Inches(2.0), Inches(5.38), Inches(3)
  graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  value_axis.maximum_scale = 100.0
  plot = chart.plots[0]
  plot.has_data_labels = True
  data_labels = plot.data_labels
  data_labels.font.size = Pt(12)
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  data_labels.position = XL_LABEL_POSITION.INSIDE_BASE
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.NONE
  category_axis.tick_labels.font.size = Pt(12)
  bar_plot = chart.plots[0]
  bar_plot.gap_width = 20
  bar_plot.overlap = -20
  value_axis.minor_tick_mark = XL_TICK_MARK.NONE
  category_axis.major_tick_mark = XL_TICK_MARK.NONE
  value_axis.major_tick_mark = XL_TICK_MARK.NONE
  category_axis.has_major_gridlines = False
  value_axis.has_major_gridlines = False
  value_axis.visible = False
  #bar_plot.fill.solid()
  #bar_plot.fill.fore_color.rgb=RGBColor(0,0,0)
  return tsg_ppt
#fills the second slide with 9 charts of the 9 questions, appears a table by the first chart with the diffrence to the last years data and with an arrow which signs the way of the difference
def fill_slide_2(tsg_ppt, q1):
  slide = tsg_ppt.slides[2]
  chart_data = ChartData()
  chart_data_2 =  ChartData()
  chart_data_3 =  ChartData()
  chart_data.categories=['1', '2']
  chart_data_2.categories=['1', '2']
  chart_data_3.categories=['1', '2', '3']
  a = [30,40,30]
  b = [20,30, 50]
  c = [20,59.4,20.6]
  d = [34, 16, 50]
  e = [10, 20, 70]
  f = [15, 75, 10]
  g = [34, 26, 40]
  chart_data.add_series('2',(a[0], b[0]))#, 0, 0))#, b[0], c[0], d[0], e[0], f[0], g[0])) #, d[0]))
  chart_data.add_series('3',(a[1], b[1]))#, 0, 0))#, b[1], c[1], d[1], e[1], f[1], g[1]))
  chart_data.add_series('4',(a[2], b[2]))#, 0, 0))#, b[2], c[2], d[2], e[2], f[2], g[2]))#, d[2]))
  chart_data_2.add_series('2',(c[0], d[0]))
  chart_data_2.add_series('3',(c[1], d[1]))
  chart_data_2.add_series('4',(c[2], d[2]))
  chart_data_3.add_series('2',(e[0], f[0], g[0]))
  chart_data_3.add_series('2',(e[1], f[1], g[1]))
  chart_data_3.add_series('2',(e[2], f[2], g[2]))
  x,y,cx,cy = Cm(10.5), Cm(4.5), Cm(5.2), Cm(3)
  graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED_100, x, y, cx, cy, chart_data)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  chart.plots[0].has_data_labels = True
  data_labels = chart.plots[0].data_labels
  bar_plot = chart.plots[0]
  bar_plot.gap_width = 10
  #nehasznaaalddata_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  chart.has_legend = False
  data_labels.font.size = Pt(12)
  data_labels.number_format = '0'
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
  category_axis.tick_labels.font.size = Pt(12)
  value_axis.minor_tick_mark = XL_TICK_MARK.NONE
  category_axis.major_tick_mark = XL_TICK_MARK.NONE
  value_axis.major_tick_mark = XL_TICK_MARK.NONE
  category_axis.has_major_gridlines = False
  value_axis.has_major_gridlines = False
  value_axis.visible = False
  category_axis.visible = False

  x,y,cx,cy = Cm(10.5), Cm(6.68), Cm(5.2), Cm(3)
  graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED_100, x, y, cx, cy, chart_data_2)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  chart.plots[0].has_data_labels = True
  data_labels = chart.plots[0].data_labels
  bar_plot = chart.plots[0]
  bar_plot.gap_width = 50
  #nehasznaaalddata_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  chart.has_legend = False
  data_labels.font.size = Pt(12)
  data_labels.number_format = '0'
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
  category_axis.tick_labels.font.size = Pt(12)
  value_axis.minor_tick_mark = XL_TICK_MARK.NONE
  category_axis.major_tick_mark = XL_TICK_MARK.NONE
  value_axis.major_tick_mark = XL_TICK_MARK.NONE
  category_axis.has_major_gridlines = False
  value_axis.has_major_gridlines = False
  value_axis.visible = False
  category_axis.visible = False

  x,y,cx,cy = Cm(10.5), Cm(8.86), Cm(5.2), Cm(4.1)
  graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED_100, x, y, cx, cy, chart_data_3)
  chart = graphic_frame.chart
  value_axis = chart.value_axis
  chart.plots[0].has_data_labels = True
  data_labels = chart.plots[0].data_labels
  bar_plot = chart.plots[0]
  bar_plot.gap_width = 50
  #nehasznaaalddata_labels.position = XL_LABEL_POSITION.OUTSIDE_END
  chart.has_legend = False
  data_labels.font.size = Pt(12)
  data_labels.number_format = '0'
  data_labels.font.color.rgb = RGBColor(0, 0, 0)
  category_axis = chart.category_axis
  category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
  category_axis.tick_labels.font.size = Pt(12)
  value_axis.minor_tick_mark = XL_TICK_MARK.NONE
  category_axis.major_tick_mark = XL_TICK_MARK.NONE
  value_axis.major_tick_mark = XL_TICK_MARK.NONE
  category_axis.has_major_gridlines = False
  value_axis.has_major_gridlines = False
  value_axis.visible = False
  category_axis.visible = False
  #chart.replace_data(chart_data)
  
  
  rows=1
  cols=1
  left = Inches(5.83)
  top = Inches(1.93)
  width = Inches(0.8)
  height = Inches (0.425)# set column widths
  table = slide.shapes.add_table(rows, cols, left, top, width, height).table
  table.columns[0].width = Inches(0.39)
  table.cell(0, 0).text = "+3"
  cell = table.rows[0].cells[0]
  paragraph = cell.textframe.paragraphs[0]
  paragraph.font.size = Pt(12)
  paragraph.font.color.rgb = RGBColor(255, 255, 255)
  cell.horizontal_anchor = MSO_ANCHOR.MIDDLE
  cell.vertical_anchor = MSO_ANCHOR.MIDDLE
  cell.fill.solid()
  cell.fill.fore_color.rgb = RGBColor(124,124,124)
  img_path = 'pirosnyil.png'
  left = Inches(6.25)
  height = Inches(0.3)
  top = Inches(2)
  pic = slide.shapes.add_picture(img_path, left, top, height=height)
  return tsg_ppt


def difference_in_shape(tsg_ppt):
  gesamt_slide = tsg_ppt.slides[2]
  rows=1
  cols=2
  left = Inches(5.83)
  top = Inches(1.93)
  width = Inches(0.8)
  height = Inches (0.425)# set column widths
  table = gesamt_slide.shapes.add_table(rows, cols, left, top, width, height).table
  table.columns[0].width = Inches(0.39)
  table.columns[1].width = Inches(0.39)
  table.cell(0, 0).text = "+3"
  table.cell(1, 0).text = "+5"
  img_path = 'pirosnyil.png'
  pleft = Inches(6.25)
  height = Inches(0.3)
  top = Inches(2)
  pic = gesamt_slide.shapes.add_picture(img_path, pleft, top, height=height)
  #table.cell(0, 1).text = pic
  cell = table.rows[0].cells[0]
  cellb = table.rows[1].cells[0]
  cell2 = table.rows[0].cells[1]
  cell2b = table.rows[1].cells[1]
  paragraph = cell.textframe.paragraphs[0]
  paragraph2 = cell2.textframe.paragraphs[0]
  paragraph.font.size = Pt(12)
  paragraph.font.color.rgb = RGBColor(255, 255, 255)
  cell.horizontal_anchor = MSO_ANCHOR.MIDDLE
  cell.vertical_anchor = MSO_ANCHOR.MIDDLE
  cell.fill.solid()
  cell2.fill.background()
  cell.fill.fore_color.rgb = RGBColor(124,124,124)
  paragraphb = cellb.textframe.paragraphs[0]
  paragraph2b = cell2b.textframe.paragraphs[0]
  paragraphb.font.size = Pt(12)
  paragraphb.font.color.rgb = RGBColor(255, 255, 255)
  #cell.horizontal_anchor = MSO_ANCHOR.MIDDLE
  cell.vertical_anchor = MSO_ANCHOR.MIDDLE
  cellb.fill.solid()
  cell2b.fill.background()
  cellb.fill.fore_color.rgb = RGBColor(124,124,124)
  
 #paragraph.alignment = PP_ALIGN.MIDDLE
  #table.cell.font.size=Pt(10.5)
  #table.cell.font.color.rgb = RGBColor(226, 0, 116)


  #table.cell(1, 0).text = '-4'
  #table.cell(1, 1).text = "nyil2"
  #table.cell(2, 0).text = '-5'
  #table.cell(2, 1).text = "nyil3"
def picture_insert(tsg_ppt, a, n, h):
  slide=tsg_ppt.slides[a]
  rows = n
  cols = 2
  left = Inches(5.9)
  top = Inches(1.93)
  width = Cm(1.72)
  height = Cm(n*h)
  table = slide.shapes.add_table(rows, cols, left, top, width, height).table
  for i in range (0, n):
    for j in range (0, 2):
      row = table.rows[i].cells[j]
      table.cell(i, 0).text = str(i)
      p = row.textframe.paragraphs[0]
      p.font.size = Pt(12)
      p.font.color.rgb = RGBColor(255,255,255)
      table.cell(i, 0).vertical_anchor = MSO_ANCHOR.MIDDLE
      table.columns[0].height = Cm(h)
      table.columns[1].height = Cm(h)
      row.vertical_anchor = MSO_ANCHOR.MIDDLE
      table.cell(i, 0).fill.solid()
      table.cell(i, 0).fill.fore_color.rgb = RGBColor(124,124,124)
      table.cell(i, 1).fill.background()
      img_path = "pirosnyil.png"
      left = Inches(6.25)
      height = Cm(h/2.05)
      if n==3:
        t = (h+2.21*i)+5.3
      print(t)
      top = Cm(t)
      pic = slide.shapes.add_picture(img_path, left, top, height=height)




############MAIN#####################  
tsg_ppt=Presentation('tsg_templ_uj.pptx')
orglongname = "Verkauf Deutschland"
orgshortname = "VLD"
org_names = ["TSG", "GF OG", "VLD", "RVL N", "RVL O", "RVL W", "RVL M", "RVL SW", "RVL S"] #tomb1
org_percents = ("41", "40", "40", "27", "53", "37", "28", "50", "44")
q1 = ("20", "30", "50")
n = 3
h = 1.6
picture_insert(tsg_ppt, 3, n, h)
fill_slide_0_and_titles(tsg_ppt, orglongname)
fill_slide_1(tsg_ppt, org_names, org_percents)  
fill_slide_2(tsg_ppt, q1)
tsg_ppt.save("outteszt/"+"Abkurzung_der_OrgEinheit_TSG_Leadership_Survey"+".pptx")

