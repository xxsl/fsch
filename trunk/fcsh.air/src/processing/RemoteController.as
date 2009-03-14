package processing {
    import amf.vo.CommandVO;
    import amf.vo.DataVO;
    import amf.vo.ErrorVO;

    import flash.events.EventDispatcher;

    import mx.logging.ILogger;
    import mx.logging.Log;

    public class RemoteController extends EventDispatcher
    {
        private static var log:ILogger = Log.getLogger("RemoteController");
        private static var _instance:RemoteController;


        public static function get instance():RemoteController
        {
            if (!_instance)
            {
                _instance = new RemoteController();
            }
            return _instance;
        }

        public function RemoteController()
        {
            super();
        }

        public function process(object:Object):void
        {
            if (object is ErrorVO)
            {
                log.info(ErrorVO(object).toString());
                dispatchEvent(new RemoteEvent(RemoteEvent.ERROR_EVENT, null, null, ErrorVO(object)));
            }
            else if (object is DataVO)
            {
                log.info(DataVO(object).toString());
                dispatchEvent(new RemoteEvent(RemoteEvent.DATA_EVENT, DataVO(object), null, null));
            }
            else if (object is CommandVO)
                {
                    log.info(CommandVO(object).toString());
                    dispatchEvent(new RemoteEvent(RemoteEvent.COMMAND_EVENT, null, CommandVO(object), null));
                }
                else
                {
                    log.error("Object can not be parsed");
                }
        }

    }
}