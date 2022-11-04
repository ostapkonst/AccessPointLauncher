# AccessPointLauncher

Скрипт для управления Wifi-точкой доступа в Windows начиная с Vista.
Из особенностей: автоматический старт точки доступа при запуске системы, автоматическая настройка службы Internet Connection Sharing.

## Установка

```
AccessPointLauncher -mode=install -run_installed=yes
AccessPointLauncher -mode=shared
```

## Удаление

```
AccessPointLauncher -mode=uninstall
```

## Команды

Полное описание доступных команд:

```
Usage: AccessPointLauncher [-mode] [-pause_timeout] [-check_admin]
            [-wait_services] [-check_hostednetwork] [-paused_exit]
            [-logging] [-latest_log] [-run_installed] [-clear_log]

Default: AccessPointLauncher -mode=help -pause_timeout=20
                             -check_admin=yes -wait_services=yes
                             -check_hostednetwork=yes -paused_exit=no
                             -logging=no -latest_log=yes -run_installed=no
                             -clear_log=no

Options:
    -mode=[help|status|start|stop|fresh|restart|install|uninstall|shared|logs]
    -pause_timeout=<positive number>
    -check_admin=[yes|no]
    -wait_services=[yes|no]
    -check_hostednetwork=[yes|no]
    -paused_exit=[yes|no]
    -logging=[yes|no]
    -latest_log=[yes|no]
    -run_installed=[yes|no]
    -clear_log=[yes|no]
```
