from ldap3 import Server, Connection, ALL, NTLM
from ldap3 import ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES, AUTO_BIND_NO_TLS, SUBTREE
import subprocess


def user_add(userInfo):
    # define the server
    # define an unsecure LDAP server, requesting info on DSE and schema
    s = Server('LDAP://hogehoge',
               use_ssl=True, get_info=ALL)

    # define the connection
    try:
        c = Connection(s, user='foofoo',
                       password='fugafuga', authentication=NTLM, auto_bind=True)
        print(c.extend.standard.who_am_i())
        c.add(
            # ここのcnが一覧に出てくる名前
            'cn={},ou=NewUsers,dc=hoge,dc=fuga,dc=foo,dc=co,dc=jp'.format(
                userInfo['cn']),
            attributes={
                'objectClass': [
                    'user', 'person', 'organizationalPerson'
                ],
                # プロパティで設定する部分
                # 姓
                'sn': userInfo['sn'],
                # 名
                'givenName': userInfo['givenName'],
                # 表示名
                'displayName': userInfo['displayName'],
                # ユーザログオン名(2000より前)
                'sAMAccountName': userInfo['sAMAccountName'],
                # ユーザログオン名
                'userPrincipalName': userInfo['userPrincipalName'],
                # 説明
                'description': userInfo['description'],
                # パスワード無期限設定
                'userAccountControl': 66048
            }
        )
        print(c.result)
        # close the connection
        c.unbind()
    except Exception as e:
        print(e)


def user_search():
    search_base = 'DC=hoge,DC=fuga,DC=foo,DC=co,DC=jp',
    server = Server('LDAP://hogehoge', get_info=ALL)
    conn = Connection(server, user='fugafuga',
                      password='foofoo', authentication=NTLM, auto_bind=True)
    print(conn.extend.standard.who_am_i())
    conn.search(search_base,
                '(&(|(objectclass=user)(objectclass=person)(objectclass=inetOrgPerson)(objectclass=organizationalPerson))(!(objectclass=computer)))',
                attributes=[ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES])

    for entry in sorted(conn.entries[:-15]):
        try:
            print(entry.sAMAccountName, entry.name)
            print(entry)
        except Exception as e:
            print("error: {}".format(e))
