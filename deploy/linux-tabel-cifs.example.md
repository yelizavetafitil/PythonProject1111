# Табель на Linux: кодировка + тот же путь `\\srv-doc\...`

В **настройках портала пути не меняют** — остаются `\\srv-doc\ТАБЕЛЬ` как на Windows.
Модуль `tabel_fs.py` на Linux подставляет mount `/mnt/tabel` и чинит кириллицу в именах файлов.

## Один раз на сервере (не ручное копирование табелей)

```bash
sudo apt install -y cifs-utils
sudo mkdir -p /mnt/tabel
# /etc/smb-tabel.creds — username, password, domain (chmod 600)

sudo mount -t cifs //srv-doc/ТАБЕЛЬ /mnt/tabel \
  -o credentials=/etc/smb-tabel.creds,uid=www-data,gid=www-data,iocharset=utf8

sudo -u www-data ls /mnt/tabel
```

`iocharset=utf8` — важно для папок **ОЦ**, **ТАБЕЛЬ** без «иероглифов».

В fstab — см. прежний пример с `//srv-doc/ТАБЕЛЬ` → `/mnt/tabel`.

## Переменные портала

**Не обязательно** менять `TABEL_BASE_DIR` — можно оставить Windows-путь или не задавать (дефолт `\\srv-doc\ТАБЕЛЬ`).

Опционально, если mount не в `/mnt/tabel`:

```ini
Environment=TABEL_LINUX_DEFAULT_BASE=/mnt/tabel
```

## Автообновление

Портал сам сканирует шару ~раз в 3 минуты (`tabel_portal_cache.json`).  
Файлы с srv-doc **не копируют вручную** — читаются через mount.

## Проверка

```bash
sudo systemctl restart portal
```

`/api/tabel/meta`: `base_dir_exists: true`, `resolved_base_dir: "/mnt/tabel"`.
