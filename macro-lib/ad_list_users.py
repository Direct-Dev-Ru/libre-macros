# -*- coding: utf-8 -*-
"""
Вывод пользователей Active Directory из указанной OU.

Подключение — аргументы макроса или строка 2 листа «AD»:
  A2 сервер, B2 пользователь, D2 auth (SIMPLE/NTLM),
  E2 SSL (1/0), F2 порт, G2 OU.
  Пароль — в диалоге при запуске (поле со *), на лист не пишется.

Пример строки 2 (лист «AD»):
  A2  dc01.corp.local
  B2  reader@corp.local
  D2  SIMPLE
  E2  1                           (Windows: LDAPS 636)
  F2  636
  G2  OU=Sales,DC=corp,DC=local
      или путь: Sales/corp/local  →  OU=Sales,DC=corp,DC=local
      или: name1/name2/dc1/c/dc2  (три DC в конце пути)

  ad_list_users("Sales/corp/local")
  ad_list_users("OU=Sales,DC=corp,DC=local")

Результат — строка 4 (заголовки), данные с строки 5.
"""
MACRO_VERSION = "3.10.706"
from libre_macros_ad_lib import ad_list_users as _ad_list_users


def ad_list_users(ou_dn=u"", server=u"", bind_user=u"", bind_password=None, auth=u"", use_ssl=None, port=None, sheet_name=u""):
    return _ad_list_users(
        ou_dn=ou_dn,
        server=server,
        bind_user=bind_user,
        bind_password=bind_password,
        auth=auth,
        use_ssl=use_ssl,
        port=port,
        sheet_name=sheet_name,
    )


g_exportedScripts = (ad_list_users,)
