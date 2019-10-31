Задание:

У нас есть таблицы с курсами валют и справочник валют, которые используются в системе. 
Таблица с курсами валют называется "Курс", в ней поля: дата, валюта, курс. 
Пример записи: ’08.21.2019’, ‘EUR’, 73,9766. 
Мы забираем данные с сайта ЦБ и заполняем эту таблицу каждый день новым курсом. 
С сайта ЦБ берем такой строчкой: http://www.cbr.ru/scripts/XML_daily.asp?date_req=21.08.2019 . 
Справочник валют называется "Валюты", в нем есть поле валюта, кодвалюты. 
Пример записи: ’RUR’, ’R00000’; ’CZK’, ’R01760’. 
Надо написать код , который будет получать курс валюты, которая у нас используется в системе, каждый день и сохранять его в таблицу "Курс" базы данных. 
Функцию, которая будет возвращать курс валюты на определенную дату. 
Функция должна принимать два параметра – валюту и дату курса, а возвращать курс. 
