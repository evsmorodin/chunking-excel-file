# chunking-excel-file
VBA-скрипт для деления файла Excel на части с сохранением структуры таблицы. 

В скрипте указаны комментарии и его довольно просто адаптивать под свои нужды.

При работе скрипта с активного листа Excel вырезается количество строк, указанное в параметре `Limit`.

Предполагается, что данные начинаются со второй строки, в первой - заголовок таблицы

### Как использовать 
Импортировать скрипт в Excel и запустить:

* Нажать `ALT`+`F11` в Excel
* Выбрать пункт меню `File` -> `Import File...`
* Выбрать файл скрипта
* Выбрать пункт меню `Run` -> `Run Macro` или нажать `F5` на клавиатуре
