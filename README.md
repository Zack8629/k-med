# Парсер Excel файлов для медицинской компании:
Предназначен для парсинга Excel файлов со списками застрахованных лиц.
Компилируется в exe файл, после чего используется компанией под Windows.

### Инструкция по применению:
Полученные данные записываются в файл для записи: "готовый файл.xlsm".
При отсутствии данного файла в момент запуска парсера он будет создан автоматически.
Все остальные файлы не обязательны к наличию — будут пропущены.
Названия файлов и папок не привязаны к регистру.
Файлы для парсинга не привязаны ни к регистру, ни к расширению.

Для расширения определения гендера можно заносить дополнительные имена
в соответствующие текстовые файлы (с большой буквы(!), через запятую/пробел/перенесение строки).

Парсер работает по всем Excel файлам, находящимся в папках, соответствующих названиям компаний:
- "Альфа"
- "Альянс"
- "Ингосстрах"
- "Ренессанс"
- "Ресо"
- "Росгосстрах"
- "Согаз"
- "Согласие 13" - (файлы у которых первая запись клиента начинается на 13 строке)
- "Согласие 15" - (файлы у которых первая запись клиента начинается на 15 строке)

После прочтения файла парсером он (файл) перемещается во вложенную папку "прочитанные файлы"
(создается автоматически в случае отсутствия).

Схема расположения файлов:
- "парсер"/"готовый файл.xlsm"
- "парсер"/списки имён/"мужские(женские) имена.txt"
- "парсер"/списки от СК/"папка из списка"/"файлы"
- "парсер"/списки от СК/"папка из списка"/прочитанные файлы/"прочитанные файлы"
