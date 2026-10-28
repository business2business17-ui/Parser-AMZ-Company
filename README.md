# Parser-AMZ-Company — Amazon Seller Intelligence

## Исправление прерванных запусков

### Страница проверки Amazon вместо профиля

Amazon может вернуть HTTP 200 и перенаправить запрос на `/checkpoint/continue`. Это страница проверки, а не пустой профиль продавца. Теперь она помечается `CHECKPOINT_REQUIRED`; запуск без окна сохраняет прогресс и завершается с `ACTION_REQUIRED`. Ошибочная страница и её адрес/заголовок сохраняются локально в `diagnostics` внутри папки запуска. Эти файлы не отправляются в GitHub.

Для продолжения в PowerShell остановите старый процесс через Ctrl+C и запустите с видимым браузером, указав папку последнего запуска:

```powershell
& "S:/Чехлы/Claude/Parser/run.bat" --no-headless --resume "reports/20261028_192240_789930_UTC"
```

Если Amazon покажет проверку, завершите её вручную в открывшемся браузере. Скрипт ждёт до 180 секунд и затем повторно открывает исходную страницу. При незавершённой проверке прогресс сохраняется; автоматическое решение CAPTCHA или смена сессий для обхода проверки не выполняются. Если профиль не распознан по другой причине, предусмотрено ожидание содержимого и один повтор. После трёх неуспешных профилей подряд запуск останавливается с диагностикой, а не продолжает создавать пустые отчёты.

Браузер по умолчанию работает без отдельного окна (`headless: true`), чтобы случайное закрытие окна не прерывало сбор. Для видимого окна используйте `--no-headless`. При закрытии или падении браузера обработка останавливается, сохраняет `checkpoint.json` и `result.xlsx` со статусом `INTERRUPTED`. Остальные ASIN получают `NOT_PROCESSED`; они не считаются обработанными.

Профиль продавца загружается сразу после первого товара этого продавца. Прогресс сохраняется после каждого товара, общий XLSX — после профиля продавца, каждых 10 товаров и при остановке. Статусы запуска: `COMPLETED`, `PARTIAL`, `INTERRUPTED`, `FAILED`.

Продолжить ваш прерванный запуск на Windows:

```bat
run.bat --resume "reports/20261028_074516_194206_UTC"
```

Укажите нужную папку запуска вместо примера. Поддерживаются и старые папки с `result.xlsx` и JSON продавцов, и новые папки с `checkpoint.json`. Скрипт использует уже полученные товары и заново запрашивает неуспешные карточки и профили. Создаётся новая папка результата; исходная сохраняется. Время наблюдения у ранее полученных товаров остаётся прежним.

Для проверки на небольшой выборке:

```bat
run.bat --limit 5 --skip-storefront
```

`--skip-storefront` собирает только входные карточки и их продавцов без расширения каталога. Лист `Seller Feedback` общего XLSX содержит отдельные оценки и количества за 1/3/12 месяцев и lifetime. Значения берутся из блоков соответствующего периода, а не из округлённой иконки звёзд. Если продавцом указан Amazon, сохраняется его отображённое название и статус `RETAIL_PROFILE_NOT_APPLICABLE`; юридические данные стороннего продавца не выдумываются.

**Цель:** от ASIN/URL или выгрузки Xray перейти к продавцу, собрать видимые сведения о карточке, seller profile и товарах на его storefront и создать **отдельный XLSX для каждого (`marketplace`, `seller_id`)**.

## Файлы

- `agent.py` — запускаемый парсер (Playwright + BeautifulSoup + openpyxl).
- `seller_profile.py` — извлечение юридических сведений, контактов и рейтингов из блоков продавца.
- `requirements.txt` — зависимости.
- `seeds_example.csv` — пример входа.
- `DEMO_SUPFINE_US.xlsx` — **иллюстративный образец**, заполненный только значениями из присланных скриншотов. Не результат онлайн-парсинга.

## Поддерживаемые маркетплейсы

| Код | Домен | Код | Домен |
|---|---|---|---|
| US | amazon.com | CA | amazon.ca |
| MX | amazon.com.mx | UK | amazon.co.uk |
| DE | amazon.de | FR | amazon.fr |
| ES | amazon.es | IT | amazon.it |
| PL | amazon.pl | NL | amazon.nl |
| BE | amazon.com.be | SE | amazon.se |
| AE | amazon.ae | SA | amazon.sa |
| IN | amazon.in | AU | amazon.com.au |
| JP | amazon.co.jp | | |

`MARKETS` в исходнике задаёт домен и Marketplace ID для каждого рынка. **Seller ID следует воспринимать как идентификатор внутри marketplace.** Достоверность Marketplace ID при производственном запуске сверяйте с документацией Amazon.

## Переносимый запуск

Скрипт рассчитан на запуск из обычной папки на Windows или Linux. На каждом компьютере нужен Python 3.11 или новее. Скопируйте репозиторий, положите одну или несколько выгрузок Helium 10 Xray (`.xlsx`) в папку `xray helium10` и запустите:

Windows: `setup.bat` выполняется один раз, затем `run.bat`. Linux/macOS: `./setup.sh`, затем `./run.sh`.

Каждый запуск создаёт отдельную папку `reports/YYYYMMDD_HHMMSS_ffffff_UTC/`. В ней находится общий файл `result.xlsx`, отдельные файлы по продавцам, JSON и `run.log`. Общий XLSX создаётся даже если Amazon недоступен: строки семян сохраняются со статусом причины (`BLOCKED_CHALLENGE`, `ERROR:...`, `NO_SELLER_ID`, `NOT_PROCESSED`).

В `Products` сохраняются количество отзывов листинга, оценка товара, исходный текст `Ships from` и классификация доставки (`AMAZON`, `SELLER`, `SHOWN_TEXT_ONLY`, `UNKNOWN`). В `Seller Profile` сохраняются количество отзывов/оценка продавца, публичные email и телефон, юридическое имя, адрес, номер регистрации, VAT, EORI, ссылки на публичные юридические документы и статусы `contact_status` / `registration_status`.

Если поле не показано Amazon, оно остаётся пустым с отметкой `NOT_SHOWN`. Если страница закрыта CAPTCHA или ограничением, в `Run Log` и `page_status` записывается `BLOCKED_CHALLENGE` или технический статус ошибки. Парсер не обходит CAPTCHA, аутентификацию или другие ограничения доступа.

Параметры по умолчанию находятся в `config.json`:

```json
{
  "input_dir": "xray helium10",
  "output_dir": "reports",
  "marketplace": null,
  "headless": true,
  "max_pages": 3,
  "max_products": 100,
  "delay": 2.0
}
```

Маркетплейс определяется из Amazon URL, колонки `Marketplace`, имени файла (`Xray_US.xlsx`) или `--marketplace US`. Скрипт больше не задаёт вопрос в консоли, поэтому одинаково работает в CI и на другом ПК. Проверка входных выгрузок без открытия браузера:

```bash
python agent.py --validate-only --marketplace US
```

## Установка вручную

Python 3.11+; установить браузер Playwright:

```bash
cd amazon_seller_agent
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
python -m playwright install chromium
```

Один товар:

```bash
python agent.py --asin B0D9VS4PV7 --marketplace US --out reports
```

По Amazon URL:

```bash
python agent.py --url 'https://www.amazon.com/dp/B0D9VS4PV7' --out reports
```

Helium 10 Xray Excel или CSV с колонкой `ASIN` / `URL` (либо `Product URL`, `link`, `product_url`):

```bash
python agent.py --input xray_export.xlsx --out reports --max-pages 5 --max-products 300
```

Для одного файла в папке проекта:

```bash
python agent.py --input "xray helium10" --marketplace US
```

## GitHub Actions

В репозитории добавлен workflow `.github/workflows/xray.yml`. Загрузите Xray `.xlsx` в `xray helium10`, закоммитьте файл, откройте GitHub → **Actions → Process Xray → Run workflow**. После выполнения скачайте артефакт `xray-reports-...`: внутри будет `result.xlsx`, отдельные отчёты и журнал. Workflow автоматически ставит Python, зависимости и Chromium и запускает браузер в headless-режиме.

Выгрузки Xray могут содержать чувствительные коммерческие данные. Если репозиторий публичный, добавьте папку `xray helium10` в `.gitignore` или используйте приватный репозиторий. Не коммитьте логин, пароль или cookies Amazon.

ASIN в гиперссылках Xray также поддерживается. При защите страницы, CAPTCHA и ограничении доступа программа помечает `BLOCKED_CHALLENGE`. По умолчанию браузер без окна; `--no-headless` открывает видимое окно. Выполняйте только разрешённый для вашего использования сбор данных.

## Как устроена цепочка

```text
Xray XLSX/CSV / ASIN / URL
     ↓ нормализация marketplace, ASIN
Товар /dp/ASIN
     ↓ sold_by / seller_id
Профиль /sp? ... seller=ID
     ↓ seller data + feedback при обнаружении фильтра
Storefront /s?me=ID&marketplaceID=...
     ↓ обнаруженные ASIN → подробные страницы товаров
1 Excel + 1 JSON на (marketplace, seller ID)
```

По EAN и бренду **ещё нет работающего автоматического резолвера**: эти поля должны подключаться через источник каталожного поиска, например Amazon SP-API Catalog Items при наличии авторизованного доступа. Не следует автоматически выбирать первый найденный товар: нужна проверка `EAN → ASIN`, `brand → несколько ASIN` и маркировка неоднозначностей. Для Xray XLSX кликабельный ASIN читается при условии, что его текстом является сам ASIN/URL.

## Схема XLSX

1. `Dashboard`: seller, marketplace, страна, число **обнаруженных** карточек, low-rating, in-stock, дата сбора.
2. `Products`: ASIN, ссылки на товар и **первую картинку (URL)**, title, bullet points/highlights, устройство, цена (raw/numeric), валюта, Buy Box, наличие, ships from, sold by, вариации, число предложений, рейтинг, отзывы, BSR, first available, статус, время.
3. `Seller Profile`: идентификатор, ссылки seller/storefront, legal entity, юридический адрес, ISO-код страны, видимый рейтинг, 12m positive, возвраты; отдельные поля для 1/3/12 мес и lifetime.
4. `Item Details`: произвольные характеристики `ASIN / Field / Value`, а не только заданный список.
5. `Run Log`: URL/время/статус обращения.

Красная заливка всей строки `Products`, если `Product Rating` — число **строго меньше 4.1**. Пустое значение = **неизвестно**, а не ноль.

### Принципы качества данных

- Товарные рейтинги и отзывы (`Customer Reviews`) **не смешивать** с seller feedback.
- Если продавец открылся, но юридический адрес не найден, `country_code` оставлять пустым. **Не выводить страну из названия магазина или Ships from.**
- `In Stock = UNKNOWN`, если Amazon не дал однозначного статуса. Доступность зависит от региона доставки.
- `Buy Box = UNKNOWN` — эвристика MVP; в производственной версии надо сверять Featured Offer и продавца offer, а отсутствие цены не равно `NO`.
- `Offer Count` в MVP извлекается из видимого текста; отсутствие счётчика ≠ один продавец. Число уникальных seller offers и число продавцов *не всегда совпадают*.
- Видимая дата `Date First Available` **не всегда** равна фактической дате создания листинга. Храним её под точным названием Amazon.
- Вариации: `YES/NO/UNKNOWN` — не делать категоричный вывод при динамически загружаемом twister.
- Storefront: `/s?me=...` **не гарантирует точного полного списка**: выдача может иметь ограничения, лимиты страниц, скрытые варианты, рекламу и нестабильный total. Значение в dashboard — `Observed products`, а не «всего активных листингов».
- Если в выдаче storefront открывается карточка с Buy Box другого продавца, отмечается `DIFFERENT_BUYBOX_SELLER` — не исключать ASIN автоматически.
- Seller feedback за 1/3/12/lifetime заполняется только когда действительно удалось обнаружить и выбрать соответствующий вариант интерфейса; иначе `NOT_COLLECTED`. Не подставлять общий рейтинг во все 4 периода.

## Что уже работает, и что добавить для production

| Модуль | Состояние |
|---|---|
| 16 доменов и marketplace IDs | конфигурация создана |
| Ввод ASIN/URL/CSV/Xray XLSX | реализовано для явных текстовых ASIN/URL |
| Product HTML parsing | реализованы базовые селекторы; нуждается в тестах для локализаций |
| Seller HTML parsing | частично: EN-селекторы/текст, не гарантирует все страны |
| Seller feedback 1/3/12/lifetime | экспериментально: чтение dropdown, значения могут отсутствовать |
| Storefront и пагинация | базовая пагинация, нет гарантии полноты |
| Один XLSX + JSON на продавца | реализовано |
| Alarm < 4.1 | реализовано через Excel conditional formatting |
| EAN/Brand discovery | интерфейс для будущего расширения; в этой версии нет |
| Режим полного production-мониторинга | не реализован |

**Следующая инженерная итерация:** API-провайдеры для каталога/offer, точный seller-offer parser, учёт A/B HTML и 16 локализаций, извлечение `num sellers`, верификация полноты storefront, отдельный слой storage (`SQLite`), повторные прогоны с дедупликацией, тестовый корпус HTML, очередь заданий, лимиты запросов и более надёжная диагностика. Отдельный режим browser-assisted для разрешённых страниц; не обходить CAPTCHA, блокировки и ограничения доступа.

### Источники

- Amazon Marketplace URLs: https://docs.business.amazon.com/docs/amazon-business-urls
- Amazon SP-API Catalog Items: https://developer-docs.amazon.com/sp-api/lang-en_EN/docs/manage-product-listings-guide
- Amazon terms (пример UK): https://www.amazon.co.uk/gp/help/customer/display.html?nodeId=GLSBYFE9MGKKQXXM

**Важно:** MVP не был проверен на живом Amazon в данной среде. Автоматические результаты будут зависеть от доступности страниц и выбранного marketplace; `DEMO` — не подтверждение успешного онлайн-запуска.
