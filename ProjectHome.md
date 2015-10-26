# [Flex Compiler SHell](http://labs.adobe.com/wiki/index.php/Flex_Compiler_Shell) integration with [Apache Ant](http://ant.apache.org/) #
Current [version 1.0.133](http://code.google.com/p/fsch/wiki/changelog?ts=1244485613&updated=changelog)

### [Post your feature requests here](http://code.google.com/p/fsch/issues/list) ###


# How it works #
  1. Download and install FCSHServer.
  1. Set environment variable FCSHServer={application directory, e.g. C:\Program File\Fcsh Server\}.
  1. Edit [server.ini](http://code.google.com/p/fsch/wiki/Preferences) file: setup Flex SDK location (e.g. sdk=C:\Flex\_sdk\_3.0) or modify Flex Compiler SHell vmoptions

Application directory contains fcsh.jar file. So example ant build.xml is:
```
<?xml version="1.0" encoding="UTF-8"?>
<project name="project.main" basedir="." default="build">
    <property environment="env"/>
    <taskdef name="fcsh" classname="fcsh">
        <classpath>
            <pathelement location="${env.FCSHServer}/fcsh.jar"/>
        </classpath>
    </taskdef>

    <target name="build">
        <fcsh consoleencoding="cp866">
            <arg value="mxmlc"/>
            <arg value="-output=C:\target.swf"/>
            <arg value="-load-config+=C:\work\FLX\src\flex-config.xml"/>
        </fcsh>
    </target>
</project>
```

fcsh has optional attribute **consoleencoding** (useful when error message text is not in English), default value is "cp866" for Cyrillic. [Supported enodings](http://java.sun.com/j2se/1.5.0/docs/guide/intl/encoding.doc.html). Also see [useful macros](http://code.google.com/p/fsch/wiki/EliminatingDependencies).

Ant tries to connect to the FCSHServer (localhost:40000), on fail it tries to launch FCSHserver again, if connection fails after 5 retries BuildException is thrown.

All subsequent builds will reuse compiler cache.

```
c:\work\google.code\fcsh.ant\test>ant
Buildfile: build.xml

build:
     [fcsh] Server is not responding. Probably it is stopped. Trying to launch...
     [fcsh] Server started
     [fcsh] Trying to connect... Attempt 0 of 5
     [fcsh] Server is up!
     [fcsh] Command: mxmlc -locale en_US -output=C:\realworld.swf -load-config+=C:\work\realworld\FLX\src\flex-config.xml
     [fcsh] fcsh: Assigned 1 as the compile target id
     [fcsh] Loading configuration file C:\work\3.3\frameworks\flex-config.xml
     [fcsh] Loading configuration file C:\work\realworld\FLX\src\flex-config.xml
     [fcsh] C:\realworld.swf (1072748 bytes)
     [fcsh] (fcsh)

     [fcsh] Awesome!

BUILD SUCCESSFUL
Total time: 8 seconds
c:\work\google.code\fcsh.ant\test>ant
Buildfile: build.xml

build:
     [fcsh] Command: mxmlc -locale en_US -output=C:\realworld.swf -load-config+=C:\work\realworld\FLX\src\flex-config.xml
     [fcsh] Loading configuration file C:\work\3.3\frameworks\flex-config.xml
     [fcsh] Loading configuration file C:\work\realworld\FLX\src\flex-config.xml
     [fcsh] Nothing has changed since the last compile. Skip...
     [fcsh] C:\realworld.swf (1072743 bytes)
     [fcsh] (fcsh)

     [fcsh] Awesome!

BUILD SUCCESSFUL
Total time: 1 second
c:\work\google.code\fcsh.ant\test>
```

FCSHServer adds tray icon, right click to see menu:

> ![http://fsch.googlecode.com/files/tray_menu.png](http://fsch.googlecode.com/files/tray_menu.png)

  * **About**, shows status, version info
> ![http://fsch.googlecode.com/files/about.png](http://fsch.googlecode.com/files/about.png)
  * **Compiler cache**, displays window where you can clean targets or recompile them manually
> http://fsch.googlecode.com/files/fcshserver_1.PNG
  * **View log**, shows log file
  * **Exit**, stops server

### Statistics ###
&lt;wiki:gadget url="http://www.ohloh.net/p/19645/widgets/project\_basic\_stats.xml" height="220" border="1"/&gt;