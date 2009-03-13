package processing {
    import amf.vo.CommandVO;
    import amf.vo.DataVO;
    import amf.vo.ErrorVO;

    import flash.events.EventDispatcher;

    public class RemoteController extends EventDispatcher
    {
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
                fcsh.instance.log(ErrorVO(object).description);
                dispatchEvent(new RemoteEvent(RemoteEvent.ERROR_EVENT, null, null, ErrorVO(object)));
            }
            else if (object is DataVO)
            {
                fcsh.instance.log(DataVO(object).toString());
                dispatchEvent(new RemoteEvent(RemoteEvent.DATA_EVENT, DataVO(object), null, null));
            }
            else if (object is CommandVO)
                {
                    fcsh.instance.log(CommandVO(object).toString());
                    dispatchEvent(new RemoteEvent(RemoteEvent.COMMAND_EVENT, null, CommandVO(object), null));
                }
                else
                {
                    fcsh.instance.log("Object can not be parsed");
                }
        }

    }
}