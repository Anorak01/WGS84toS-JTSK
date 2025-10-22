run:
	. venv/bin/activate
	pip install -r requirements.txt
	python3 main.py

prep:
	python3 -m venv venv

package:
	( \
	rm -rf tempvenv main.spec build; \
	python3 -m venv tempvenv; \
	source ./tempvenv/bin/activate; \
	which pip; \
	which python3; \
	pip install -r requirements.txt; \
	pip install pyinstaller; \
	python3 -m PyInstaller --onefile --add-data "./table_yx_3_v1710.dat:." --add-data "./CR-2005_v1005.dat:." --noconsole main.py; \
	rm -rf tempvenv main.spec build; \
	)
