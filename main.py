#coding=utf-8
from selenium import webdriver
from selenium.webdriver.support.ui import Select

import time
import os
import sys
from datetime import datetime
import traceback
import glob

import smtplib
from email.mime.text import MIMEText

# global variables
csv_dir = './csv_report/'

def login(browser,user,passwd):
	browser.find_element_by_id('USER').send_keys(user)
	browser.find_element_by_id('PASSWORD').send_keys(passwd)
	browser.find_element_by_id('safeLoginbtn').click()
	time.sleep(3)

def get_vah_group_info(browser,vah_group_name):
	global csv_dir
	with open(csv_dir + vah_group_name + '_group.csv','w') as f:
		f.write('VAAName,Status,Partition Count\n')
	browser.find_element_by_xpath('//li[@id="' + vah_group_name + '"]/a').click()
	time.sleep(3)

	# one group has 4 vah
	for i in range(1,5):
		name = browser.find_element_by_xpath("//*[@class='table table-striped table-bordered center']" \
			"/tbody/tr[" + str(i)+ "]/td[1]/a").text
		status = browser.find_element_by_xpath("//*[@class='table table-striped table-bordered center']" \
			"/tbody/tr[" + str(i) + "]/td[3]").text
		par_cnt = browser.find_element_by_xpath("//*[@class='table table-striped table-bordered center']" \
			"/tbody/tr[" + str(i) + "]/td[5]").text

		with open(csv_dir + vah_group_name + '_group.csv','a') as f:
			f.write(name + ',' + status + ',' + par_cnt + '\n')

	with open(csv_dir + vah_group_name + '_group.csv','r') as f:
		lines = f.readlines()
		del lines[0]
		if lines[0].split(',')[1]  == 'Live' and lines[1].split(',')[1] == 'Standby' and lines[2].split(',')[1] == 'Standby' and lines[3].split(',')[1] == 'Standby':
			print vah_group_name + ' status ok'
		else:
			tmp = vah_group_name + ' Live/Stanby status abnormal\n'
			with open('errors.txt','a') as f:
				f.write(tmp)

def get_vah_vaa_info(browser,vah_group_name,vah_name):
	global csv_dir
	with open(csv_dir + vah_group_name + '.' + vah_name + '_vaa_info.csv','w') as f:
		f.write('PNum,Product,VAA,State\n')
	vaa_info_list = []

	lines = int(browser.find_element_by_xpath("//div[@id='vaalist']").get_attribute('pnum'))
	for i in range(1,lines+1):
		pnum = browser.find_element_by_xpath("//div[@id='vaalist']/table/tbody/tr[" + str(i) + "]/td[2]").text
		product = browser.find_element_by_xpath("//div[@id='vaalist']/table/tbody/tr[" + str(i) + "]/td[3]").text
		# ignore empty lines
		if product == '':
			continue

		vaa = browser.find_element_by_xpath("//div[@id='vaalist']/table/tbody/tr[" + str(i) + "]/td[4]").text
		state = browser.find_element_by_xpath("//div[@id='vaalist']/table/tbody/tr[" + str(i) + "]/td[7]/span").text
		if state != 'Activated':
			with open('errors.txt','a') as f:
				f.write(vah_group_name + '.' + vah_name + ': partition' + pnum + ' not in Activated state\n')

		tmp_list = pnum + ',' + product + ',' + vaa
		vaa_info_list.append(tmp_list)

		with open(csv_dir + vah_group_name + '.' + vah_name + '_vaa_info.csv','a') as f:
			f.write(pnum + ',' + product + ',' + vaa + ',' + state + '\n')
	return vaa_info_list

def __get_detailed_rule_info(rule_name,product_name):
	global csv_dir
	dict = {}
	rules_info = open(csv_dir + product_name + '_liveRuleList.csv').readlines()
	for rule_info in rules_info:
		# remove \n
		rule_info = rule_info[0:-1]
		rule_list = rule_info.split(',')
		if rule_name.upper() == rule_list[0].upper():
			dict['name'] = rule_name
			dict['de#'] = rule_list[1]
			dict['owner'] = rule_list[2]
			dict['created_at'] = rule_list[3]
			dict['updated_at'] = rule_list[4]
			return dict
	return dict

def get_vah_vaa_statistics_info(browser,vah_group_name,vah_name,vaa_info_list):
	global csv_dir

	with open(csv_dir + vah_group_name + '.' + vah_name + '_vaa_statistics.csv','w') as f:
		f.write('Partition Name,Num of Rules,Num of RICs subscribed,Num of RICs created,Rules List\n')

	f2 = open(csv_dir + vah_group_name + '_rules.csv','w')
	f2.write('Rule,PNum,Product,VAA,Owner,Created At,Updated At\n')
	#,Num of RICs subscribed,Num of RICs created

	select = Select(browser.find_element_by_xpath("//select[@id='select-interpreter']"))

	options = select.options
	print vah_group_name + '.' + vah_name + ':'
	for i in range(1,len(options)):
		select.select_by_index(i)
		time.sleep(3)
		par_name = select.all_selected_options[0].text
		# partition1-appshell is not our own, so ignore it
		if par_name == 'partition1-appshell':
			continue
		content = par_name

		rule_sel = Select(browser.find_element_by_xpath("//select[@id='select-rule']"))

		rule_list = ''
		# only activaited partitions have 0-n rules. Loaded partitions don't have
		if len(rule_sel.options) <= 1:
			num_of_rules = 0
			num_of_rics_subscribed = 0
			num_of_rics_created = 0
		else:
			for j in range(1,len(rule_sel.options)):
				# get rule list of one partition
				rule_sel.select_by_index(j)
				rule_name = rule_sel.all_selected_options[0].text
				rule_list = rule_list + '|' + rule_name

				for i in range(0,len(vaa_info_list)):
					tmp_list = vaa_info_list[i].split(',')
					if ('partition' + tmp_list[0]) in par_name:
						break
				dict = __get_detailed_rule_info(rule_name,tmp_list[1])
				if dict:
					f2.write(rule_name + ',' + tmp_list[0] + ',' + tmp_list[1] + ',' + \
		                    tmp_list[2] + ',' + dict['owner'] + ',' + dict['created_at'] +
					         ',' + dict['updated_at'] + '\n')
				else:
					f2.write(rule_name + ',' + tmp_list[0] + ',' + tmp_list[1] + ',' + \
		                    tmp_list[2] + '\n')

		btn_vaa_statistics = browser.find_element_by_xpath('//a[@id="btn-confirmInterpreter"]')

		btn_vaa_statistics.click()
		# proper delay is important
		time.sleep(4)

		'''
		try:
			rule_input = browser.find_element_by_xpath('//table[@id="entryTab"]/tbody[1]/tr/td[3]').text
			rule_output = browser.find_element_by_xpath('//table[@id="entryTab"]/tbody[2]/tr/td[3]').text
		except:
			pass
		'''

		num_of_rules = browser.find_element_by_xpath("//*[@id='appTab']/tbody[1]/tr/td[3]").text
		num_of_rics_subscribed = browser.find_element_by_xpath("//*[@id='appTab']/tbody[2]/tr/td[3]").text
		num_of_rics_created = browser.find_element_by_xpath("//*[@id='appTab']/tbody[3]/tr/td[3]").text

		content = content + ',' + num_of_rules + ',' + num_of_rics_subscribed + ',' + num_of_rics_created
		print '[INFO]',content
		content = content + ',' + rule_list[1:] + '\n'
		with open(csv_dir + vah_group_name + '.' + vah_name + '_vaa_statistics.csv','a') as f:
			f.write(content)

def get_vah_info(browser,vah_group_name,vah_name):
	time.sleep(2)
	browser.find_element_by_xpath('//*[@id="' + vah_group_name + '.' + vah_name + '"]/a').click()
	time.sleep(3)
	# after clicking vah element or general tab, vaa information can be got
	vaa_info_list = get_vah_vaa_info(browser,vah_group_name,vah_name)

    # status: Live or Standby
	#status = browser.find_element_by_xpath("//div[@id='vahstatus']/table/tbody/tr[3]/td[2]").text

	browser.find_element_by_xpath("//*[@id='tabs']/ul/li[3]/a").click()
	time.sleep(2)
	# after clicking VSS Statistics, vss-statistics information can be got
	get_vah_vaa_statistics_info(browser,vah_group_name,vah_name,vaa_info_list)

def __get_rule_general_info(browser,rule_line_num):
	dic = {}
	created_at = browser.find_element_by_xpath('//div[@id="general-page"]/table[1]/tbody/tr[6]/td[2]').text
	updated_at = browser.find_element_by_xpath('//div[@id="general-page"]/table[2]/tbody/tr[1]/td[2]/a').text
	name = browser.find_element_by_xpath('//div[@id="general-page"]/table[1]/tbody/tr[1]/td[2]/a').text
	de = browser.find_element_by_xpath('//div[@id="general-page"]/table[1]/tbody/tr[7]/td[2]').text
	owner = browser.find_element_by_xpath('//div[@id="general-page"]/table[1]/tbody/tr[4]/td[2]').text

	dic['name'] = name
	dic['de#'] = de
	dic['owner'] = owner
	dic['created_at'] = created_at
	dic['updated_at'] = updated_at

	return dic

def get_live_rules(browser,product_name):
	''' get rule list from /VDI/Rule	'''
	global csv_dir
	rule_state_sel = Select(browser.find_element_by_xpath("//*[@id='rule-state']"))
	# select 'Has Activated Version' option
	rule_state_sel.select_by_index(3)
	browser.find_element_by_xpath('//div[@id="search-criteria"]/a[@data-action="search"]').click()
	time.sleep(3)

	rule_num = browser.find_element_by_xpath('//div[@class="span8"]/div/span').text
	idx = 0

	rule_list = []
	with open(csv_dir + product_name + '_liveRuleList.csv','w') as f:
		f.write('Rule Name,DE#,Rule Owner,Created At,Updated At\n')
	page_tab = browser.find_elements_by_xpath('//div[@id="paginagtion"]/ul/li')

	print 'Live Rule List for ' + '"' + product_name + '"(%s):' %(rule_num)
	for page in range(2,len(page_tab)-2):
		ele_rule_list = browser.find_elements_by_xpath('//table[@id="rule-list"]/tbody/tr')
		for rule in range(1,len(ele_rule_list)+1):
			try:
				browser.find_element_by_xpath('//table[@id="rule-list"]/tbody/tr[' + \
			                                      str(rule) + ']/td[1]/a')
			except Exception,e:
				# if label a doesn't exist, then all live rules have been visited
				# But this kind of statement isn't good, since more time consumed
				return rule_list

			d = {}
			browser.find_element_by_xpath('//table[@id="rule-list"]/tbody/tr[' + str(rule) + ']/td[3]').click()
			d = __get_rule_general_info(browser,rule)
			print d
			rule_name = d['name']
			rule_de = d['de#']
			rule_owner = d['owner']
			rule_created_at = d['created_at']
			rule_updated_at = d['updated_at']

			rule_list.append(d)

			idx = idx + 1

			with open(csv_dir + product_name + '_liveRuleList.csv','a') as f:
				f.write(rule_name + ',' + rule_de + ',' + rule_owner + ',' + rule_created_at + ',' + rule_updated_at + '\n')
		# click next page
		if not (page == len(page_tab)-3):
			browser.find_element_by_xpath('//div[@id="paginagtion"]/ul/li[' + str(page+2) + ']/a').click()
			time.sleep(3)
	if rule_num != idx:
		print '[ERROR] rule num error'

def __get_servers_rule_list(product_name,vah_group_names):
	'''  get rule list from /VAH/Servers '''
	global csv_dir
	rule_list = []
	for j in range(0,len(vah_group_names)):
		vaa_lines = open(csv_dir + vah_group_names[j] + '.UK1P-' + vah_group_names[j] + 'A_vaa_info.csv').readlines()
		del vaa_lines[0]
		for line in vaa_lines:
			vaa_line_record = line.split(',')
			product = vaa_line_record[1]
			if product == product_name:
				pnum = vaa_line_record[0]
				par_lines = open(csv_dir + vah_group_names[j] + '.UK1P-' + vah_group_names[j] + 'A_vaa_statistics.csv').readlines()
				del par_lines[0]
				for par_line in par_lines:
					par_line_list = par_line.split(',',4)
					if not ('partition' + pnum) in par_line_list[0]:
						continue
					tmp = par_line_list[4]
					# delete \n
					tmp = tmp[0:len(tmp)-1]
					tmp_list = tmp.split('|')
					#print 'vah_group_name:%s, partition%s, tmp_list(%d):%s' %(vah_group_names[j],pnum,len(tmp_list),tmp_list)
					for i,val in enumerate(tmp_list):
						rule_list.append(val)
					break
	#print '\nproduct_name-rule_list(%d):%s' %(len(rule_list),rule_list)
	return rule_list

def __compare_rule(product_name,vah_group_names,rule_list):
	'''  get rule list from /VAH/Servers '''
	global csv_dir
	for j in range(0,len(vah_group_names)):
		vaa_lines = open(csv_dir + vah_group_names[j] + '.UK1P-' + vah_group_names[j] + 'A_vaa_info.csv').readlines()
		del vaa_lines[0]
		for line in vaa_lines:
			vaa_line_record = line.split(',')
			product = vaa_line_record[1]
			if product == product_name:
				pnum = vaa_line_record[0]
				par_lines = open(csv_dir + vah_group_names[j] + '.UK1P-' + vah_group_names[j] + 'A_vaa_statistics.csv').readlines()
				del par_lines[0]
				for par_line in par_lines:
					par_line_list = par_line.split(',',4)
					if not ('partition' + pnum) in par_line_list[0]:
						continue
					tmp = par_line_list[4]
					# delete \n
					tmp = tmp[0:len(tmp)-1]
					tmp_list = tmp.split('|')
					#print 'vah_group_name:%s, partition%s, tmp_list(%d):%s' %(vah_group_names[j],pnum,len(tmp_list),tmp_list)
					for i,val in enumerate(tmp_list):
						if val == '' or val == 'de42979':
							continue
						if rule_in_list(val,rule_list) == True:
							#print '[Rule] ' + val + ' consistent'
							pass
						else:
							with open('errors.txt','a') as f:
								f.write('Live rule <font color="red">' + val + '</font> in machine ' + vah_group_names[j] +
								    ':partition' + pnum + ' not exists in ' + product_name + ' Rule list\n')
					break

def rule_in_list(rule,rule_list):
	for k,vdi_rule in enumerate(rule_list):
		if rule.upper() == vdi_rule.upper():
			return True
	return False

def get_error_status(product_list,vah_group_names):
	'''
	f = open('Abnormal.csv','w')
	f.write('Rule Name,Asset Class,ERROR Type,ERROR Detail\n')
	f.close() # if not closed, compare_rule_list() write operation will suffer
	'''
	global csv_dir

	for i in range(0,len(product_list)):
		vdi_rule_list = []
		rule_lines = open(csv_dir + product_list[i] + '_liveRuleList.csv','r').readlines()
		del rule_lines[0]
		for line in rule_lines:
			rule_name = line.split(',')[0]
			vdi_rule_list.append(rule_name)
		#print '\n' + product_list[i] + '  --- VDI Rule list(%d):%s' %(len(vdi_rule_list),vdi_rule_list)

		__compare_rule(product_list[i],vah_group_names,vdi_rule_list)

def process(browser,vah_group_name):
	# generate 4 groups's files named *_group.csv
	try:
		get_vah_group_info(browser,vah_group_name)
	except:
		pass
	time.sleep(1)
	browser.find_element_by_xpath('//li[@id="' + vah_group_name + '"]/ins').click()
	time.sleep(1)

	name1pa = 'UK1P-' + vah_group_name + 'A'
	name1pb = 'UK1P-' + vah_group_name + 'B'
	name2pa = 'UK2P-' + vah_group_name + 'A'
	name2pb = 'UK2P-' + vah_group_name + 'B'

	get_vah_info(browser,vah_group_name,name1pa)
	#get_vah_info(browser,vah_group_name,name1pb)
	#get_vah_info(browser,vah_group_name,name2pa)
	#get_vah_info(browser,vah_group_name,name2pb)

def __write_errors(f):
	lines = open('errors.txt','r').readlines()
	#del lines[len(lines)-1]

	if len(lines) == 0:
		str = '<h2 id="abnormal_rules">NO Errors</h2><hr />'
		f.write(str)
	else:
		str = '<h2 id="abnormal_rules">Errors</h2>'
		f.write(str)

		str = '<ul>'
		for line in lines:
			print '[Error] ',line
			str = str + '<li>' + line + '</li>'
		str = str + '</ul></br><hr />'
		f.write(str)

	'''
	lines = open('Abnormal.csv','r').readlines()
	# delete table header
	del lines[0]
	# show by order
	lines.sort()

	str = '\n<table border="1" align="center">\n' \
	    '<tr><th>Rule Name</th>' \
		'<th>Asset Class</th>' \
		'<th>Error Type</th>' \
		'<th>Error Detail</th></tr>\n'
	for line in lines:
		tmp_list = line.split(',')
		str = str + '<tr>' \
			'<td>' + tmp_list[0] + '</td>' \
			'<td>' + tmp_list[1] + '</td>' \
			'<td>' + tmp_list[2] + '</td>' \
			'<td>' + tmp_list[3] + '</td>' \
			'</tr>\n'
	f.write(str + '</table></br><hr />')
	'''

def __write_changes(f,vah_group_names):
	global csv_dir
	tmp_str = '<h2 id="changes">Rule Changes</h2><table border="1" align="center">' \
				'<tr><th>Operation</th><th>Product</th><th>VAH-PNum</th><th>Rule</th></tr>'

	flag = 0
	for p in vah_group_names:
		new_lines = open(csv_dir + p + '_rules.csv').readlines()
		old_lines = open(csv_dir + p + '_rules_old.csv').readlines()
		for line in new_lines:
			if line in old_lines:
				old_lines.remove(line)
			else:
				new_eles = line.split(',')

				tmp_str += '<tr><td>Add</td><td>' + new_eles[2] + '</td>' \
				          '<td>' + p + '-P' + new_eles[1] + '</td>' \
				          '<td><font color="red">' + new_eles[0] + '</td></tr>'
				flag = 1
		if len(old_lines) != 0:
			for i in old_lines:
				old_eles = i.split(',')
				tmp_str += '<tr><td>Delete</td><td>' + old_eles[2] + '</td>' \
				          '<td>' + p + '-P' + old_eles[1] + '</td>' \
				          '<td><font color="red">' + old_eles[0] + '</td></tr>'
			flag = 1
	if flag != 0:
		f.write(tmp_str + '</table></br><hr />')
	else:
		f.write('<h2 id="changes">NO Rule Changes</h2><hr />')

def __write_live_rules(f,product_list):
	global csv_dir
	tmp_str = '<h2 id="live_rules">Live Rules</h2>\n<table border="1" align="center">\n' \
		'<tr><th>Product</th><th>Num of Rules</th></tr>\n'
	f.write(tmp_str)
	# export live rules info
	for i in range(0,len(product_list)):
		live_rule_str = open(csv_dir + product_list[i] + '_liveRuleList.csv').readlines()
		new_rule_lines = len(live_rule_str)-1

		if os.path.exists(csv_dir + product_list[i] + '_liveRuleList_old.csv'):
			old_live_rule_str = open(csv_dir + product_list[i] + '_liveRuleList_old.csv').readlines()
			old_rule_lines = len(old_live_rule_str)-1

		if os.path.exists(csv_dir + product_list[i] + '_liveRuleList_old.csv') and new_rule_lines != old_rule_lines:
			if new_rule_lines > old_rule_lines:
				diff = '+' + str(new_rule_lines-old_rule_lines)
			else:
				diff = str(new_rule_lines-old_rule_lines)
			tmp_str = '<tr><td>' + product_list[i] + '</td><td><font color="red">' + str(new_rule_lines) + \
			          '</font>(<font color="red">' + diff + '</font>)' + '</td><tr>\n'
		else:
			tmp_str = '<tr><td>' + product_list[i] + '</td><td>' + str(new_rule_lines) + '</td><tr>\n'

		f.write(tmp_str)

	f.write('</table></br><hr />')

def __write_load_of_servers(f,vah_group_names):
	global csv_dir
	tmp_str = '<h2 id="load_of_vah_servers">Load of VAH Servers</h2>'
	f.write(tmp_str)

	for i in range(0,len(vah_group_names)):
		tmp_str = '\n<table border="1" align="center">\n<tr><th colspan="4"><h4>' + vah_group_names[i] + '</h4></th></tr>\n' \
		'<tr><th>Partition Name</th>' \
		'<th>Num of Rules</th>' \
		'<th>Num of Source RIC</th>' \
		'<th>Num of Output RIC</th></tr>\n'
		f.write(tmp_str)

		lines = open(csv_dir + vah_group_names[i] + '.UK1P-' + vah_group_names[i] + 'A_vaa_statistics.csv').readlines()
		del lines[0]

		if os.path.exists(csv_dir + vah_group_names[i] + '.UK1P-' + vah_group_names[i] + 'A_vaa_statistics_old.csv'):
			old_lines = open(csv_dir + vah_group_names[i] +
			                 '.UK1P-' + vah_group_names[i] + 'A_vaa_statistics_old.csv').readlines()
			del old_lines[0]

		tmp_str = ''
		sum_of_rules = 0
		sum_of_subcribed = 0
		sum_of_created = 0
		for j in range(0,len(lines)):
			tmp_list = lines[j].split(',',4)

			if os.path.exists(csv_dir + vah_group_names[i] + '.UK1P-' + vah_group_names[i] + 'A_vaa_statistics_old.csv'):
				old_tmp_list = old_lines[j].split(',',4)

				tmp_str = tmp_str + '<tr>' \
				'<td align="center">' + tmp_list[0].split('-')[0] + '</td>'
				if int(tmp_list[1]) > int(old_tmp_list[1]):
					tmp_str += '<td align="center"><font color="red">' + tmp_list[1] + '</font>(' \
					           '<font color="red">+' + str(int(tmp_list[1]) - int(old_tmp_list[1])) + '</font>)</td>'
				elif int(tmp_list[1]) < int(old_tmp_list[1]):
					tmp_str += '<td align="center"><font color="red">' + tmp_list[1] + '</font>(' \
					           '<font color="red">' + str(int(tmp_list[1]) - int(old_tmp_list[1])) + '</font>)</td>'
				else:
					tmp_str += '<td align="center">' + tmp_list[1] + '</td>'

				if int(tmp_list[2]) > int(old_tmp_list[2]):
					tmp_str += '<td align="center"><font color="red">' + tmp_list[2] + '</font>(' \
					           '<font color="red">+' + str(int(tmp_list[2]) - int(old_tmp_list[2])) + '</font>)</td>'
				elif int(tmp_list[2]) < int(old_tmp_list[2]):
					tmp_str += '<td align="center"><font color="red">' + tmp_list[2] + '</font>(' \
					           '<font color="red">' + str(int(tmp_list[2]) - int(old_tmp_list[2])) + '</font>)</td>'
				else:
					tmp_str += '<td align="center">' + tmp_list[2] + '</td>'

				if int(tmp_list[3]) > int(old_tmp_list[3]):
					tmp_str += '<td align="center"><font color="red">' + tmp_list[3] + '</font>(' \
					           '<font color="red">+' + str(int(tmp_list[3]) - int(old_tmp_list[3])) + '</font>)</td>'
				elif int(tmp_list[3]) < int(old_tmp_list[3]):
					tmp_str += '<td align="center"><font color="red">' + tmp_list[3] + '</font>(' \
					           '<font color="red">' + str(int(tmp_list[3]) - int(old_tmp_list[3])) + '</font>)</td>'
				else:
					tmp_str += '<td align="center">' + tmp_list[3] + '</td></tr>'

			else:
				tmp_str = tmp_str + '<tr>' \
				'<td align="center">' + tmp_list[0].split('-')[0] + '</td>' \
				'<td align="center">' + tmp_list[1] + '</td>' \
				'<td align="center">' + tmp_list[2] + '</td>' \
				'<td align="center">' + tmp_list[3] + '</td>' \
				'</tr>\n'

			sum_of_rules = sum_of_rules + int(tmp_list[1])
			sum_of_subcribed = sum_of_subcribed + int(tmp_list[2])
			sum_of_created = sum_of_created + int(tmp_list[3])

		tmp_str = tmp_str + '<tr><td align="center">sum</td>' \
		    '<td align="center">' + str(sum_of_rules) + '</td>' \
			'<td align="center">' + str(sum_of_subcribed) + '</td>' \
			'<td align="center">' + str(sum_of_created) + '</td>' \
			'</tr>\n'
		f.write(tmp_str + '</table></br>')

def to_html(filename,product_list,vah_group_names):
	str = '<html><head><title>Rule Info</title></head><body>' \
		'<h1>Rule Statistics</h1>\n' \
		'<div align="right"><p>Report generated at:' + time.strftime("%Y-%m-%d") + ' ' + time.strftime('%H:%M:%S') + \
		'\n</div><div class="nav"><ul>' \
	    '<li><a href="#abnormal_rules">Errors</a></li>\n' \
		'<li><a href="#changes">Rule Changes</a></li>\n' \
		'<li><a href="#live_rules">Live Rules</a></li>\n' \
		'<li><a href="#load_of_vah_servers">Load of VAH Servers</a></li></div>' \
		'<hr />\n'
	f = open(filename,'w')
	f.write(str)

	__write_errors(f)
	__write_changes(f,vah_group_names)
	__write_live_rules(f,product_list)
	__write_load_of_servers(f,vah_group_names)

	str = '</body></html>'
	f.write(str)

def send_email_with_other_program():
	os.system('EmailSender.exe')

def rename_files(srcdir):
	oldfiles = glob.glob('csv_report\\*_old.csv')
	for file in oldfiles:
		os.remove(file)
	files = os.listdir(srcdir)
	for file in files:
		dirname = os.getcwd() + '\\csv_report'
		name = os.path.splitext(file)[0]
		srcfile = dirname + '\\' + name + '.csv'
		dstfile = dirname + '\\' + name + '_old.csv'
		os.rename(srcfile,dstfile)

def send_exceptions_email(sender,content):
	host = '164.179.13.19'
	receivers = []
	receivers.append(sender)

	msg = MIMEText(content,'plain','utf-8')
	msg['Subject'] = 'Exception...'

	mail = smtplib.SMTP(host,25)
	mail.sendmail(sender,receivers,msg.as_string())
	mail.quit()

if __name__ == '__main__':
	main_url = ''
	vah_group_names = []
	product_list = []
	product_list_capital = []
	
	cfg_info = open('config.txt').readlines()
	for line in cfg_info:
		if line == '' or line[0] == '#':
			continue
		eles = line.split('=')
		if eles[0] == 'user':
			user = eles[1][0:-1]
		elif eles[0] == 'passwd':
			# remove \n
			passwd = eles[1][0:-1]
		elif eles[0] == "sender":
			sender = eles[1][0:-1]
		elif eles[0] == 'main_url':
			main_url = eles[1][0:-1]
		elif eles[0] == 'vah':
			vah_group_names = eles[1].split(',')
			vah_group_names[len(vah_group_names)-1] = vah_group_names[len(vah_group_names)-1][0:-1]
		elif eles[0] == 'product':
			product_list = eles[1].split(',')
			product_list[len(product_list)-1] = product_list[len(product_list)-1][0:-1]
			for p in product_list:
				product_list_capital.append(p.upper())
		else:
			pass

	browser = webdriver.Chrome()
	#browser.maximize_window()
	browser.set_window_size(1000,600)
	browser.get(main_url)
	browser.implicitly_wait(30)
	login(browser,user,passwd)

	print '[TIME]',datetime.now()
	start = time.time()

	real_product_list = []
	try:
		rename_files(os.path.abspath('') + '\\csv_report')
		
		# let the product menu show
		browser.find_element_by_xpath('//div[@id="ProductSelect"]/a').click()

		# get all the product tag including what's not in config.txt
		product_list_tag = browser.find_elements_by_xpath('//div[@id="ProductSelect"]/ul/li/a')

		for t in product_list_tag:
			# record the VDI C AND E tag
			if t.text.upper() == 'VDI C AND E':
				vdi_c_and_e_tag = t
			real_product_list.append(t.text.upper())
		# the menu still there, so click 'VDI C AND E' tag, then the product menu is gone
		vdi_c_and_e_tag.click()
		time.sleep(4)
		# enter /VDI/Rule page
		browser.find_element_by_xpath("//*[@href='/VDI/Rule']").click()
	except Exception,e:
		print '[Exception]',traceback.format_exc()
		send_exceptions_email(sender,str(datetime.now()) + '\n' + traceback.format_exc())	

	# previous product_list_tag can't be used again any more !!!

	for i in range(0,len(real_product_list)):
		try:
			browser.find_element_by_xpath('//div[@id="ProductSelect"]/a').click()
			product_name = browser.find_elements_by_xpath('//div[@id="ProductSelect"]/ul/li/a')[i].text
			if product_name.upper() not in product_list_capital:
				# let the menu disappear
				browser.find_element_by_xpath('//div[@id="ProductSelect"]/a').click()
				continue
			print '\nProduct --- ' + product_name
			browser.find_elements_by_xpath('//div[@id="ProductSelect"]/ul/li/a')[i].click()
			time.sleep(3)
			get_live_rules(browser,product_name)

		except Exception,e:
			print '[Exception]',traceback.format_exc()
			send_exceptions_email(sender,str(datetime.now()) + '\n' + traceback.format_exc())

	browser.get(main_url)
	time.sleep(5)

	f = open('errors.txt','w')
	f.close()

	for i in range(0,len(vah_group_names)):
		print '\nVAH Group Host --- ' + vah_group_names[i]
		try:
			process(browser,vah_group_names[i])
		except Exception,e:
			print '[Exception]',traceback.format_exc()
			send_exceptions_email(sender,str(datetime.now()) + '\n' + traceback.format_exc())

	print '[TIME]',datetime.now()
	end = time.time()
	print 'Total Online time: ' + str((end-start)/60.0) + ' min'

	try:
		get_error_status(product_list,vah_group_names)
	except Exception,e:
		print '[Exception]',traceback.format_exc()
		send_exceptions_email(sender,str(datetime.now()) + traceback.format_exc())

	try:
		to_html('report.html',product_list,vah_group_names)
	except Exception,e:
		print '[Exception]',traceback.format_exc()
		send_exceptions_email(sender,str(datetime.now()) + '\n' + traceback.format_exc())

	send_email_with_other_program()
	browser.quit()
