---
layout: wide_main
title: Code side-by-side
nav_order: 101
permalink: /code-side-by-side
---

# Document Title

The usual [Markdown Cheatsheet](https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet)
does not cover some of the more advanced Markdown tricks, but here
is one. You can combine verbatim HTML with your Markdown. 
This is particularly useful for tables.
Notice that with **empty separating lines** we can use Markdown inside HTML:

<table>
<tr>
<th>Json 1</th>
<th>Markdown</th>
</tr>
<tr>
<td>

```json
{
  "id": 1,
  "username": "joewwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww",
  "email": "joe@example.com",
  "order_id": "3544fc0"
}
```

</td>
<td>

```json
{
  "id": 5,
  "username": "maryaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
  "email": "mary@example.com",
  "order_id": "f7177da"
}
```

</td>
</tr>
</table>