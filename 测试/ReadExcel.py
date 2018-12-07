# -*- coding:utf8 -*-
import ConfigParser

con = ConfigParser.ConfigParser()
config_path = '../../config/config.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    server = con.get('info', 'server')
    username = con.get('info', 'username')
    password = con.get('info', 'password')
    database = con.get('info', 'database')
print server, username, password, database
