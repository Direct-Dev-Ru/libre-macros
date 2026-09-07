# AD Linux CLI — список пользователей из OU

Автономный каталог для **внешнего Python на Linux**: выгрузка пользователей Active Directory (LDAPS + NTLM/SIMPLE).

Не требует LibreOffice / AlterOffice и **не ссылается на macro-lib** — bundled ldap3 лежит локально в `bundled/`.

## Состав

| Путь | Назначение |
|------|------------|
| `list_ad_users.py` | CLI: пользователи → CSV |
| `list_ad_computers.py` | CLI: компьютеры → CSV |
| `ad_cli_common.py` | Общие параметры, getpass, запись CSV |
| `ldap3_ad.py` | Подключение ldap3, paged search |
| `bundled/ldap3_bundled.py` | Локальная копия pyasn1 + ldap3 (~2,5 МБ) |
| `sync_bundles.py` | Обновить `bundled/` из macro-lib после пересборки |

Каталог `ad-linux/` можно скопировать на сервер целиком и запускать без остального репозитория.

## Требования

- Python 3.8+
- Доступ к DC по LDAP/LDAPS (обычно **636**)
- Учётная запись с правом чтения OU
- Файл `bundled/ldap3_bundled.py` (идёт в репозитории; обновление — см. ниже)

### Обновление бандла (из полного репозитория)

```bash
python3 macro-lib/bundle_ldap3.py
python3 ad-linux/sync_bundles.py
```

## Колонки результата (CSV, UTF-8 BOM)

**Пользователи** (`list_ad_users.py`):

`sAMAccountName`, `displayName`, `mail`, `title`, `department`, `telephoneNumber`, `userPrincipalName`, `groups`, `distinguishedName`.

**Компьютеры** (`list_ad_computers.py`):

`name`, `sAMAccountName`, `dNSHostName`, `operatingSystem`, `operatingSystemVersion`, `description`, `enabled` (1/0), `lastLogon`, `whenCreated`, `whenChanged`, `distinguishedName`.

По умолчанию разделитель — **запятая**; файл с `-f` пишется в **UTF-8 с BOM** (удобно открывать в Excel).

---

## Быстрый старт

> Примеры ниже: `cd ad-linux`, затем команды.

### Интерактивно

```bash
cd /path/to/ad-linux
python3 list_ad_users.py
```

Пароль вводится через **getpass** (символы не отображаются).

### Минимум параметров (пароль — getpass)

```bash
cd ad-linux
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\ldap_reader' \
  -o 'OU=Sales,DC=corp,DC=local'
```

  -f users.csv
```

---

## Компьютеры AD

### Все PC в OU (CSV в файл)

```bash
python3 list_ad_computers.py \
  -s dc01.corp.local \
  -u 'CORP\ldap_reader' \
  -o 'OU=Workstations,OU=Computers,DC=corp,DC=local' \
  -f computers.csv
```

### Вся ветка Computers

```bash
python3 list_ad_computers.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -o 'OU=Computers,DC=corp,DC=local' \
  -f all_computers.csv \
  -q
```

### NTLM + LDAPS, интерактив

```bash
python3 list_ad_computers.py -i -f computers.csv
```

### SIMPLE + вывод в stdout

```bash
python3 list_ad_computers.py \
  -s dc01.corp.local \
  -u reader@corp.local \
  -a SIMPLE \
  -o 'DC=corp,DC=local' | head -20
```

---

## Формат OU (`-o` / G2)

Поддерживаются два вида записи:

| Формат | Пример | Результат (LDAP DN) |
|--------|--------|---------------------|
| Классический DN | `OU=Sales,DC=corp,DC=local` | без изменений |
| Путь через `/` | `Sales/Users/corp/local` | `OU=Users,OU=Sales,DC=corp,DC=local` |
| Путь + 3 DC | `name1/name2/dc1/c/dc2` | `OU=name2,OU=name1,DC=dc1,DC=c,DC=dc2` |

Правила для пути:

- сегменты слева направо: **OU от корня к листу**, в конце — **метки домена (DC)**;
- число DC в конце: **авто** по `-s` (`dc01.corp.local` → 2 сегмента `corp/local`), иначе **2**;
- явно: `--dc-parts 3` для `…/dc1/c/dc2`.

```bash
# авто: хвост corp/local совпадает с доменом dc01.corp.local
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'reader@corp.local' \
  -a SIMPLE \
  -o 'Sales/Users/corp/local'

# три DC в конце пути
python3 list_ad_users.py \
  -s dc01.example.com \
  -u 'reader@example.com' \
  -a SIMPLE \
  -o 'name1/name2/dc1/c/dc2' \
  --dc-parts 3
```

Тот же формат в ячейке **G2** листа «AD» макроса Calc.

---

## Примеры: аутентификация (пользователи)

### NTLM + LDAPS (по умолчанию)

Формат пользователя: `DOMAIN\login` или `login@domain.com`.

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\ldap_reader' \
  -a NTLM \
  --ssl \
  -o 'OU=Sales,DC=corp,DC=local' \
  -f /tmp/sales_users.csv
```

Короткие ключи:

```bash
python3 list_ad_users.py \
  -s 192.168.10.10 \
  -u 'CORP\reader' \
  -o 'OU=Users,OU=Sales,DC=corp,DC=local'
```

### SIMPLE (UPN) + LDAPS

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u reader@corp.local \
  -a SIMPLE \
  --ssl \
  -o 'OU=Sales,DC=corp,DC=local'
```

### SIMPLE + DN пользователя

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CN=ldap reader,OU=Service,DC=corp,DC=local' \
  -a SIMPLE \
  --ssl \
  -o 'DC=corp,DC=local'
```

---

## Примеры: подключение

### Явный порт LDAPS 636

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -o 'OU=IT,DC=corp,DC=local' \
  --ssl \
  -P 636
```

### Plain LDAP без TLS (порт 389)

Только если домен разрешает нешифрованный LDAP.

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -o 'OU=IT,DC=corp,DC=local' \
  --no-ssl \
  -P 389
```

### Несколько DC (указать конкретный хост)

```bash
python3 list_ad_users.py \
  -s dc02.corp.local \
  -u 'CORP\reader' \
  -o 'OU=Remote,DC=corp,DC=local'
```

---

## Примеры: вывод

### CSV в stdout (пайп в другую команду)

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -o 'OU=Sales,DC=corp,DC=local' \
  -q | head -20
```

### CSV в файл (по умолчанию)

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -o 'OU=Sales,DC=corp,DC=local' \
  -f ./sales.csv
```

Файл: UTF-8 BOM, разделитель `,` (меняется через `--delimiter`).

### CSV в stdout

Сообщения о подключении и счётчик — в stderr; `-q` их отключает.

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -o 'OU=Sales,DC=corp,DC=local' \
  -f users.csv \
  -q
```

---

## Примеры: интерактив и пароль

### Дополнить недостающие поля в консоли

Заданы server и OU — user и пароль спросит:

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -o 'OU=Sales,DC=corp,DC=local'
```

### Полностью интерактивный режим (`-i`)

Даже если часть параметров передана, переспросит все поля (кроме пароля при `-p`):

```bash
python3 list_ad_users.py -i
```

### Пароль в переменной окружения (скрипты/CI)

```bash
export AD_PASS='секрет'
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -p "$AD_PASS" \
  -o 'OU=Sales,DC=corp,DC=local' \
  -q
```

> Не используйте `-p` в интерактивной истории shell — пароль попадёт в `.bash_history`.

### Cron / без TTY

Без терминала **getpass не сработает** — передайте `-p` или запускайте из wrapper с TTY.

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -p "$AD_PASS" \
  -o 'OU=Sales,DC=corp,DC=local' \
  -f /var/reports/ad_users.csv \
  -q
```

---

## Примеры: большие OU

### Увеличить размер страницы paged search

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -o 'OU=Company,DC=corp,DC=local' \
  --page-size 1000 \
  -f all_users.csv
```

---

## Примеры: использование как модуля

```python
from ldap3_ad import search_users, search_computers

rows = search_computers(
    server="dc01.corp.local",
    user=r"CORP\ldap_reader",
    password="***",
    ou_dn="OU=Computers,DC=corp,DC=local",
    auth="NTLM",
    use_ssl=True,
)
for row in rows:
    print(row["dNSHostName"], row["operatingSystem"], row["enabled"])
```

---

## Справка по ключам

```bash
python3 list_ad_users.py --help
```

| Ключ | Описание |
|------|----------|
| `-s`, `--server` | Хост DC |
| `-u`, `--user` | Bind: `DOMAIN\user` или UPN |
| `-p`, `--password` | Пароль (иначе getpass) |
| `-o`, `--ou` | Base DN или путь `OU1/OU2/corp/local` |
| `--dc-parts` | Сколько последних сегментов пути — DC (auto по `-s`, иначе 2) |
| `-a`, `--auth` | `NTLM` (default) или `SIMPLE` |
| `--ssl` | LDAPS, порт 636 (default) |
| `--no-ssl` | LDAP, порт 389 |
| `-P`, `--port` | Порт явно |
| `-f`, `--output` | Файл CSV |
| `--delimiter` | Разделитель (default `,`) |
| `--page-size` | Paged search (default 500) |
| `-i`, `--interactive` | Спросить поля в консоли |
| `-q`, `--quiet` | Без сообщений в stderr |

---

## Типичные ошибки

### `automatic bind not successful - invalidCredentials`

AD **отклонил логин или пароль** (до поиска в OU). Частые причины:

| Ситуация | Решение |
|----------|---------|
| По умолчанию `-a NTLM`, а логин `ivanov` | `-u 'CORP\ivanov'` (домен обязателен) |
| `-a NTLM`, логин `user@corp.local` | Попробуйте `-a SIMPLE -u user@corp.local` |
| `-a SIMPLE`, логин `CORP\user` | Для SIMPLE так нельзя → `-a NTLM` или UPN |
| Пароль верный, но в shell «съелся» `\` | Кавычки: `-u 'CORP\reader'`, не `-u CORP\reader` |
| На Linux NTLM капризничает | **Рекомендуется:** `-a SIMPLE -u reader@corp.local` |

**Проверка SIMPLE + UPN (часто самый простой вариант на Linux):**

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'reader@corp.local' \
  -a SIMPLE \
  --ssl \
  -o 'OU=Sales,DC=corp,DC=local' \
  -f users.csv
```

**Проверка NTLM:**

```bash
python3 list_ad_users.py \
  -s dc01.corp.local \
  -u 'CORP\reader' \
  -a NTLM \
  --ssl \
  -o 'OU=Sales,DC=corp,DC=local'
```

Убедитесь, что учётка **не заблокирована** и может **читать LDAP** (обычный пользователь домена или service account).

| Симптом | Что проверить |
|---------|----------------|
| `SSL: CERTIFICATE_VERIFY_FAILED` | CA домена, `--no-ssl` для теста, или корпоративный CA в системе |
| `Can't contact LDAP server` | Firewall 636/389, DNS имени DC, доступность `-s` |
| `Не задан --password (stdin не TTY)` | В cron/пайпе передайте `-p` или `$AD_PASS` |
| `No module named ldap3` | Нет `bundled/ldap3_bundled.py` — `python3 sync_bundles.py` |

---

## Связь с макросом Calc

| | `list_ad_users.py` | `list_ad_computers.py` | `macro-lib/ad_list_users.py` |
|--|----------------------|---------------------------|------------------------------|
| Среда | внешний Python Linux | внешний Python Linux | AlterOffice / LO |
| Объект | пользователи | компьютеры | пользователи |
| Формат | CSV | CSV | лист «AD» |
| OU | DN или `Sales/corp/local` | то же | G2: DN или путь через `/` |

Логика атрибутов и групп (`memberOf`) совпадает с макросом.
