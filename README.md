## From XLS to BD0 & VT0 files (uploading to ASV app)
1. Установить python (cmd написать python, windows сам откроет Microsoft Store и предложит установить)
3. Запустить cmd. Перейти в каталог UfkConverter и выполнить команду:
	pip install -r requirements.txt
4. Файл разместить в каталог со скриптом, имя файла обязательно должно быть VT_BD_SF.xls
5. Запустить скрипт командой:
	python ufkconverter.py
6. После выполнения скрипта в текущем каталоге будет создано 2 файла (BD0 и VT0)