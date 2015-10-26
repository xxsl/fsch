# Introduction #

Sometimes it is necessary to build project without FCSHServer installed or allow developers to choose among other build tools. This can be done via checking environment variables in Ant macros.


# Solution #

This is example of Ant macros (using [Ant-Contrib](http://ant-contrib.sourceforge.net) tasks):
```
<?xml version="1.0" encoding="UTF-8"?>
<project name="flex" basedir=".">

    <!--http://ant-contrib.sourceforge.net/-->
    <taskdef resource="net/sf/antcontrib/antcontrib.properties"/>

    <property environment="env"/>

    <property name="mxmlc.jar" location="${env.FLEX_SDK_HOME}/lib/mxmlc.jar"/>
    <property name="compc.jar" location="${env.FLEX_SDK_HOME}/lib/compc.jar"/>


    <!--FCSHServer home-->
    <property name="FCSHServer" value="${env.FCSHServer}"/>
    
    <property name="FLEX_HOME" value="${env.FLEX_SDK_HOME}"/>
    <property name="FLEX_FLEXLIB" value="${env.FLEX_SDK_HOME}/frameworks"/>
    <property name="FLEX_LOCALE" value="en_US"/>
    <property name="FLEX_HEADLESS_SERVER" value="${env.FLEX_HEADLESS_SERVER}"/>
    <property name="FLEX_DEBUG" value="true"/>
    <property name="FLEX_WARNING" value="true"/>

    <!--Optional-->
    <property name="KEEP-AS3-METADATA" value="Cachable"/>

    <macrodef name="flex">
        <attribute name="command"/>
        <element name="additional" implicit="true" optional="true"/>
        <sequential>
            <if>
                <contains string="${FCSHServer}" substring="$"/>
                <then>
                    <if>
                        <equals arg1="@{command}" arg2="mxmlc"/>
                        <then>
                            <java jar="${mxmlc.jar}" fork="true" maxmemory="512m" failonerror="true">
                                <arg value="+flexlib=${FLEX_FLEXLIB}"/>
                                <arg value="-debug=${FLEX_DEBUG}"/>
                                <arg value="-compiler.headless-server=${FLEX_HEADLESS_SERVER}"/>
                                <arg value="-locale=${FLEX_LOCALE}"/>
                                <arg value="-warnings=${FLEX_WARNING}"/>
                                <arg value="-keep-as3-metadata+=${KEEP-AS3-METADATA}"/>
                                <additional/>
                            </java>
                        </then>
                        <else>
                            <java jar="${compc.jar}" fork="true" maxmemory="512m" failonerror="true">
                                <arg value="+flexlib=${FLEX_FLEXLIB}"/>
                                <arg value="-debug=${FLEX_DEBUG}"/>
                                <arg value="-compiler.headless-server=${FLEX_HEADLESS_SERVER}"/>
                                <arg value="-locale=${FLEX_LOCALE}"/>
                                <arg value="-warnings=${FLEX_WARNING}"/>
                                <arg value="-keep-as3-metadata+=${KEEP-AS3-METADATA}"/>
                                <additional/>
                            </java>
                        </else>
                    </if>

                </then>
                <else>
                    <echo message="FCSHServer"/>
                    <taskdef name="fcshserver" classname="fcsh">
                        <classpath path="${FCSHServer}/fcsh.jar"/>
                    </taskdef>
                    <fcshserver>
                        <arg value="@{command}"/>
                        <arg value="+flexlib=${FLEX_FLEXLIB}"/>
                        <arg value="-debug=${FLEX_DEBUG}"/>
                        <arg value="-locale=${FLEX_LOCALE}"/>
                        <arg value="-compiler.headless-server=${FLEX_HEADLESS_SERVER}"/>
                        <arg value="-warnings=${FLEX_WARNING}"/>
                        <arg value="-keep-as3-metadata+=${KEEP-AS3-METADATA}"/>
                        <additional/>
                    </fcshserver>
                </else>
            </if>
        </sequential>
    </macrodef>
</project>

```

Usage:

```
        <flex command="compc">
            <arg value="-output=${frontend.swc}"/>
            <arg value="-source-path+=${projects.dir}/flex/frontend/src"/>
            <arg value="-source-path+=${projects.flex.common.resources}"/>
            <arg value="-include-sources+=${dashboard.dir}/flex/frontend/src"/>
            <arg value="-external-library-path+=${project.home}/WEB/webres/suite/libraries/frontend.swc"/>
        </flex>
```