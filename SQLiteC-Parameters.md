---
layout: default
title: Parameters
nav_order: 5
parent: SQLiteC
permalink: /sqlitec/parameters
---

The SQLiteCParameters class is responsible for parameter binding. SQLiteCParameters supports both sequential and named parameters. Any database-supported naming scheme can be used via the primary interface, *BindDictOrArray()*. It takes either a 1D array of parameter values or a dictionary, mapping \<parameter name\>&nbsp;&rarr;&nbsp;\<parameter value\>. If an array is provided, its size must match the number of placeholders. If a dictionary is provided, it may contain fewer parameters than the number of placeholders or some unrelated items. The code uses SQLite API to see if a particular key matches a named parameter, and, if yes, it binds the corresponding value. Unmatched parameters retain previously assigned values unless *BindClear()* method is called explicitly.
