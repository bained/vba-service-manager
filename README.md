# Service Manager

Настолно приложение за салон за услуги — клиенти, услуги, записвания в календар и
отчети. Работната книга е Excel с VBA, а кодът се строи и тества от скриптов хамър.

*Service Manager — a VBA/Excel desktop application for a small service business:
clients, services, appointments on a calendar and reports, built and tested by a
script harness (no add-ins required).*

## Версия

**1.1.0** — последният пред-релийз и официалният файл се генерират от хамъра (виж по-долу).

## Как се работи

| Команда | Какво прави |
|---|---|
| `cscript //nologo tools\build.vbs build` | само компилира → `build\Service_Manager.xlsm` |
| `cscript //nologo tools\build.vbs test` | + тестовете → `build\Service_Manager.test.xlsm`, `build\test-report.txt` |
| `cscript //nologo tools\build.vbs draft` | пред-релийз → `release\Service_Manager_draft_<версия>.xlsm` |
| `cscript //nologo tools\build.vbs release` | официален файл → `release\Service_Manager_<версия>.xlsm` + `README.txt` |
| `cscript //nologo tools\build.vbs lint` | само проверките на източниците |
| `cscript //nologo tools\make-template.vbs` | пре-генерира шаблона (**презаписва го**) |
| `tools\save-point.cmd "съобщение"` | комит (точка за връщане назад) |

Редът е `draft` → `release`: релийзът иска прясно `PASSED` в `build\test-report.txt`.

## Структура

- `src\app` — приложението: модули, класове и форми (кодът е на английски)
- `src\tests` — VBA тестовете + фикстурите на крос-чека (`crosscheck.txt`)
- `lib\` — вендорната библиотека (`sql_emulate`: SqlTable / SqlQuery / SqlMutation)
- `tools\` — хамърът (билд, тестове, релийз, watchdog), линтът, скриптовият модел
  (`tools\sim`, вторият свят на правилата под cscript) и помощни скриптове
- `Service_Manager.xlsx` — шаблонът: листове, имена, формати, бутони (без макроси)
- `PROJECT_CONCEPT.md` — спецификацията (единственият източник на истина); локален
  работен документ, който **не** се качва в хранилището
- `TODO.md` — фази, уроци (вкл. скъпо платените) и следващи стъпки; също локален

## Как е устроено

- **Кодът живее в `src\`.** Готовите файлове в `build\` и `release\` са генерирани:
  хамърът маха кодовите компоненти и импортва източниците при всеки билд.
- **Дизайнът живее в шаблона** `Service_Manager.xlsx`. Промяна, направена само в
  готовия `.xlsm`, се губи при следващ билд.
- **Имената са договорът.** Клетките се адресират по име (`o_*`, `s_*`, `grid_*`,
  `b_*`), а не по адрес — затова една клетка може да се мести, но не и да се
  преименува или трие (иначе името става `#REF!`).
- **Правилата се тестват в два свята**: скриптовият модел под `cscript`
  (`tools\sim`) и VBA тестовете (`src\tests`), а `crosscheck.txt` пази, че отговарят
  еднакво.

## Библиотека за заявките

Достъпът до данните не се пише на ръка: заявките към листовете (SQL-подобен език
върху таблиците) минават през **VBA емулатор на SQL заявки**, който е отделно мое
репозитори:

**<https://github.com/bained/excel-vba-sql-emulator-class>**

Библиотеката живее в `lib\` — `SqlTable.cls`, `SqlQuery.cls`, `SqlMutation.cls` — и
хамърът я импортва преди приложението (`tools\build.vbs` → `ImportModules`).
`modData` строи заявки и не обхожда клетки, така че всички правила за данните са на
едно място и се тестват и в скриптовия модел, и във VBA.

## Изтегляне

Официалният файл се публикува като **GitHub Release** — готовият `.xlsm` не е в
хранилището, то пази източниците, инструментите и шаблона:

- последна версия: **1.1.0** → `Service_Manager_1.1.0.xlsm`
- пред-релийзът за проби е `Service_Manager_draft_<версия>.xlsm`
