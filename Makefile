# Cameron Yick
# 06/01/2017
# Experimental integration with xlwings on Mac Excel 2016

# Change this to any snapshot ID
SNAPSHOT=7ff503c8-2752-4b8d-9272-91699a1214d4


quickstart:
	# xlwings quickstart myproject

install:
	pip install -r requirements.txt
	# Run first time only on Mac
	# xlwings runpython install

freeze:
	pip freeze > requirements.txt

# TESTING
test_export:
	curl -X GET 'https://public.enigma.com/api/export/$(SNAPSHOT)' \
	-H "Authorization: Bearer $(ENIGMA_API_KEY)" > demo
