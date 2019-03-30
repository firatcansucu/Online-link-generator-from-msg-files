# https://stackoverflow.com/questions/26322255/parsing-outlook-msg-files-with-python
import extract_msg
import re
import glob

text_file = open("online_links.txt", "w")
csv_file = open("online_links.csv", "w")
csv_file.write('langloc, online_link\n')
text_file.write('langloc, online_link\n')
for filename in glob.glob('*.msg'):
	msg = extract_msg.Message(filename)
	msg_subj = msg.subject
	msg_message = msg.body

	# extracting the online link
	foo = re.search("\<(.*?)\>", msg_message)
	online_link = foo.group(0)
	printed_onlinelink = online_link[1:-1] + "\n"

	# extracting the langloc from the subject
	langloc = msg_subj[5] + msg_subj[6] + msg_subj[7] + msg_subj[8] + msg_subj[9]

	# writing to files
	to_write = langloc + ',' + printed_onlinelink
	text_file.write(to_write)
	csv_file.write(to_write)
text_file.close()
csv_file.close()
