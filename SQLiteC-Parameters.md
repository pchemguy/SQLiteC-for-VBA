---
layout: default
title: Parameters
nav_order: 5
parent: SQLiteC
permalink: /sqlitec/parameters
---

The SQLiteCParameters class is responsible for parameter binding. It supports both positional anonymous and named parameters employing any database-supported naming scheme. The primary entry point, *BindDictOrArray()*, takes either a 1D array of parameter values or a dictionary, mapping \<parameter name\>&nbsp;&rarr;&nbsp;\<parameter value\>. In the former case, the size of the array must match the number of placeholders. In the latter case, the dictionary may contain fewer parameters than the number of placeholders or some unrelated items. The code uses SQLite API to see if a particular key matches a named parameter, and, if yes, it binds the corresponding value. Unmatched parameters retain previously assigned values unless reset via the *BindClear()* method.
