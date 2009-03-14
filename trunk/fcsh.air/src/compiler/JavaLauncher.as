package compiler {
    import prefs.PreferencesFacade;

    public class JavaLauncher
    {
        public static function getCommand():String
        {
            var flexSdkPath:String = PreferencesFacade.instance.getValue(PreferencesFacade.FLEX_SDK_PATH, "C:/Flex SDK 3");
            var vm_opts:String = PreferencesFacade.instance.getValue(PreferencesFacade.VM_OPTS, "-Xmx384m -Xms125m -XX:MaxPermSize=512m -Dsun.io.useCanonCaches=false -Duser.language=en");

            var result:String = "java.exe ";
            result += vm_opts;
            result += " -Dapplication.home=" + flexSdkPath;
            result += " -jar " + flexSdkPath + "/lib/fcsh.jar";
            return result;
        }
    }
}