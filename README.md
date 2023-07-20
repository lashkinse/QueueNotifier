# QueueNotifier
Программа предназначена для мониторинга предварительной записи на сайте mfc63.ru и оповещения о записи на произвольную электронную почту. Программа будет повторять отправку по таймеру, пока в списке есть хотя бы 1 клиент
![](https://github.com/lashkinse/QueueNotifier/raw/main/screens/2.png)

## Стек технологий
* autoit
* autoit-winhttp
* mailsend

# Установка
1. Скачайте со страницы релиза архив и распакуйте содержимое в произвольную папку.
2. Отредактируйте config.ini по инструкции ниже.
3. Запустите QueueNotifier.exe, при первом запуске программа проверит наличие предварительной записи на сайте и, если такая имеется, отправит на почту уведомление.

# Работа с исходниками/сборка
Для сборки необходим компилятор AutoIt. 

# Настройка
Все настройки задаются в файле config.ini. Для отправки уведомлений рекомендую создать новую почту

**RaisID** - *ID вашей площадки (см пункт RaisID ниже)*

**SmtpServer, Port** - *Параметры smtp сервера с которого будут отправляться письма (по умолчанию указаны рабочие для mail.ru)*

**FromName** - *Имя отправителя*

**FromAddress** - *Адрес отправителя*

**Login, Password** - *Логин/пароль почты отправителя*

**ToAddress** - *Куда отправлять отчет*

**AdminMail** - *Техподдержка*

**Delay** - *Частота проверок в минутах*

**DayStart** - *Начало дня*

**DayEnd** - *Конец дня*

# RaisID
Чтобы узнать свой RaisID необходимо войти на портал mfc63.samregion.ru через страницу авторизации http://mfc63.samregion.ru/user, а далее на выбор есть два способа:

- **Простой способ**

Нужный номер указан на странице http://mfc63.samregion.ru/user
![](https://github.com/lashkinse/QueueNotifier/raw/main/screens/3.png)

- **Сложный способ**

1. Перейти в предварительную запись.
2. Нажать новые заявления.
3. После того, как сформируется список заявлений (либо система скажет, что "В таблице отсутствуют данные") нажимаем правой кнопкой мыши на кнопку "Новые заявления" и выбираем "Просмотреть код"
4. Ищем "admin-request-nav" (Ctrl+F). Под четвертым результатом будут строки var id = 0; var rais = 1; rais - искомый номер, его указываем в настройках RaisID=1
