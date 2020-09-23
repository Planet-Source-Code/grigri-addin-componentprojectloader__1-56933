ComponentProjectLoader
~~~~~~~~~~~~~~~~~~~~~~

By grigri, 2004

Normally when you open a non-project VB-associated file such as a form (.frm), class module (.cls) or standard module (.bas), 
VB creates a new project and adds that component to it. I've never found this functionality useful, and on occasion it's annoying.

This Add-In detects when components are loaded directly, and tries to find the associated project.
If it succeeds, it loads the project and then displays the component you loaded - a much more useful way of doing things!