run:
	./venv/bin/python3 ./src/index.py

setup:
	python3 -m venv venv
	./venv/bin/pip install -r requirements.txt

clean:
	rm -rf __pycache__
	rm -rf venv