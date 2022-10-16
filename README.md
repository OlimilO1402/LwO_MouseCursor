# LwO_MouseCursor  
## lightweight objects, mouse cursor hour glass  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/LwO_MouseCursor?style=plastic)](https://github.com/OlimilO1402/LwO_MouseCursor/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/LwO_MouseCursor?style=plastic)](https://github.com/OlimilO1402/LwO_MouseCursor/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/LwO_MouseCursor/total.svg)](https://github.com/OlimilO1402/LwO_MouseCursor/releases/download/v1.0.2/LwOMouseCursor_v1.0.2.zip)
[![Follow](https://img.shields.io/github/followers/OlimilO1402.svg?style=social&label=Follow&maxAge=2592000)](https://github.com/OlimilO1402/LwO_MouseCursor/watchers)

Project started around summer 2006.  
COM manages the lifecycle of an object with the reference counting mechanism. Therefore every class in COM has to implement minimum the IUnknown Interface.  
IUnknown has these 3 Functions:  
* QueryInterface  
* AddRef  
* Release  
  
VB6/VBC/VBA is a perfect member of COM. So every VB-class implements the IUnknown interface.  
But moreover in VB every class also implements the IDispatch interface which again has 4 addtional functions that allow the functions of an object to be called by name and in late binding.  
IDispath has these 4 functions:  
* GetIDsOfNames  
* GetTypeInfo  
* GetTypeInfoCount  
* Invoke  
  
VB can of course deal with classes that implement the IUnknown interface only, the so called lightweight objects.  
This repo is a small example for explaining how to build a lightweight object, and how to use and run it in VB. This technic was first published by Matthew Curland in 1999  
Have a look at the book:

[http://powervb.mvps.org/](http://powervb.mvps.org/)  

![<AppName> Image](Resources/<AppName>.png "<AppName> Image")
