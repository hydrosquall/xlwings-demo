# Cameron Yick
# 06/01/2017
# Experimental integration with XLwings on Mac 2016

install:
	pip install -r requirements.txt
	# Run first time only on Mac
	# xlwings runpython install

freeze:
	pip freeze > requirements.txt
