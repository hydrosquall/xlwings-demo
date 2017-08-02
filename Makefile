# Cameron Yick
# 06/01/2017
# Experimental integration with XLwings on Mac 2016

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
curl_api:
	curl -X GET 'https://public.enigma.com/api/export/$(SNAPSHOT)' \
	-H "Authorization: Bearer $(ENIGMA_API_KEY)" > demo
