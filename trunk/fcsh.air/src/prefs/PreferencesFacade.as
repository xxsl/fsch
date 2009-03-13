package prefs {

    public class PreferencesFacade extends Preference
    {
        public static const HOST:String = "HOST";
        public static const PORT:String = "PORT";
        public static const FLEX_SDK_PATH:String = "FLEX_SDK_PATH";
        public static const TARGETS:String = "TARGETS";
        public static const FILENAME:String = "FILENAME";
        public static const VM_OPTS:String = "VM_OPTS";

        private static var _instance:PreferencesFacade;


        public static function get instance():PreferencesFacade
        {
            if (!_instance)
            {
                _instance = new PreferencesFacade();
            }
            return _instance;
        }

        public function PreferencesFacade()
        {
            super("preferences.amf");
        }
    }
}