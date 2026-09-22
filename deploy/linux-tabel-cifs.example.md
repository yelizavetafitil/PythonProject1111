# Табель на Linux: кодировка + тот же путь `\\srv-doc\...`

В **настройках портала пути не меняют** — остаются `\\srv-doc\ТАБЕЛЬ` как на Windows.
Модуль `tabel_fs.py` на Linux подставляет mount `/mnt/tabel` и чинит кириллицу в именах файлов.

## Один раз на сервере (не ручное копирование табелей)

```bash
sudo apt install -y cifs-utils
sudo mkdir -p /mnt/tabel
# /etc/smb-tabel.creds — username, password, domain (chmod 600)

sudo mount -t cifs //srv-doc/ТАБЕЛЬ /mnt/tabel \
  -o credentials=/etc/smb-tabel.creds,iocharset=utf8,vers=3.0,ro,file_mode=0644,dir_mode=0755

mountpoint /mnt/tabel
ls -la /mnt/tabel
```

`iocharset=utf8` — важно для папок **ОЦ**, **ТАБЕЛЬ** без «иероглифов».

Чтобы подключение восстанавливалось после перезагрузки, добавьте в `/etc/fstab`:

```fstab
//srv-doc/ТАБЕЛЬ /mnt/tabel cifs credentials=/etc/smb-tabel.creds,iocharset=utf8,vers=3.0,ro,file_mode=0644,dir_mode=0755,_netdev,nofail,x-systemd.automount 0 0
```

Затем примените и проверьте:

```bash
sudo systemctl daemon-reload
sudo mount -a
mountpoint /mnt/tabel
ls -la /mnt/tabel
```

## Переменные портала

**Не обязательно** менять `TABEL_BASE_DIR` — можно оставить Windows-путь или не задавать (дефолт `\\srv-doc\ТАБЕЛЬ`).

Опционально, если mount не в `/mnt/tabel`:

```ini
TABEL_HOST_MOUNT=/mnt/tabel
TABEL_REQUIRE_MOUNT=1
```

## Автообновление

Портал сам сканирует шару ~раз в 3 минуты (`tabel_portal_cache.json`).  
Файлы с srv-doc **не копируют вручную** — читаются через mount.

## Проверка

```bash
docker compose --env-file .env.production up -d --force-recreate web
docker compose --env-file .env.production exec web ls -la /mnt/tabel
```

`/api/tabel/meta`: `base_dir_exists: true`, `resolved_base_dir: "/mnt/tabel"`.
