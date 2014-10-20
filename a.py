#coding=utf8

from PIL import Image, ImageEnhance, ImageDraw, ImageFont, ImageFilter
import os


nsimsun = ImageFont.truetype('fonts/simsun.ttc', 24, 1)
dotum = ImageFont.truetype('fonts/gulim.ttc', 26, 3)
simhei = ImageFont.truetype('fonts/simhei.ttf', 24)
arial = ImageFont.truetype('fonts/arial.ttf', 24)

im = Image.open("templates/1.png")

draw = ImageDraw.Draw(im)

draw.text((370,530), u'郭建辉', font=simhei, fill='#000')
draw.text((590,530), u'mmjia81', font=arial, fill='#000')
draw.text((370,590), u'300000000048', font=arial, fill='#000')
draw.text((660,650), u'一级代理', font=simhei, fill='#000')
draw.text((254,710), u'芙源品牌一级代理经销商授权。', font=simhei, fill='#000')
draw.text((320,1080), u'2014年7月23日 - 2015年6月23日', font=simhei, fill='#000')

im.show()
im.save("output/1.png")


im = Image.open("templates/2.png")

draw = ImageDraw.Draw(im)

draw.text((410,500), u'구오', font=dotum, fill='#000')
draw.text((650,500), u'mmjia81', font=arial, fill='#000')
draw.text((400,558), u'300000000048', font=arial, fill='#000')
draw.text((680,620), u'에이전트', font=dotum, fill='#000')
draw.text((270,680), u'푸소스브랜드딜러수준의권한을부여。', font=dotum, fill='#000')
draw.text((300,1080), u'2014년7월23일 - 2015년6월23일', font=dotum, fill='#000')

im.show()
im.save("output/2.png")

