package compiler {

    import flash.events.EventDispatcher;

    public class BuildFacade extends EventDispatcher
    {
        private static var _instance:BuildFacade;

        private var _targetIDs:Object = new Object();
        private var _currentTarget:String;
        private var _output:String = "";

        public static function get instance():BuildFacade
        {
            if (!_instance)
            {
                _instance = new BuildFacade();
            }
            return _instance;
        }

        public function BuildFacade()
        {
            super();
        }


        public function get targetIDs():Object
        {
            return _targetIDs;
        }

        public function clear():void
        {
            _targetIDs = new Object();
            _output = "";
        }

        public function reset(target:String = null):void
        {
            _output = "";
            _currentTarget = target;
        }

        public function process(data:String):void
        {
            _output += data;

            var match:Array = _output.match(/\(fcsh\)/im);
            var result:Boolean = match != null && match.length > 0;
            if (result)
            {
                processResult();
            }
        }

        private function processResult():void
        {
            var errors:Array = hasErrors();
            if (!errors)
            {
                //fcsh: Assigned 1 as the compile target id
                var match:Array = _output.match(/fcsh: Assigned ([0-9]+) as the compile target id/im);
                var result:Boolean = match != null && match.length > 1;
                if (result)
                {
                    if (_currentTarget)
                    {
                        _targetIDs[_currentTarget] = match[1];
                    }
                }

                var warnings:Array = hasWarnings();

                if (warnings)
                {
                    dispatchEvent(new BuildEvent(BuildEvent.BUILD_WARNING, null, warnings));
                }
                else
                {
                    dispatchEvent(new BuildEvent(BuildEvent.BUILD_SUCCESSFULL, (result && _currentTarget != null) ? match : null));
                }
            }
            else
            {
                dispatchEvent(new BuildEvent(BuildEvent.BUILD_ERROR, null, null, errors));
            }
            reset();
        }

        private function hasErrors():Array
        {
            var match:Array = _output.match(/(.*Error:[^\r]+)/gim);
            var result:Boolean = match != null && match.length > 0;
            return result ? match: null;
        }

        private function hasWarnings():Array
        {
            var match:Array = _output.match(/(.*Warning:[^\r]+)/gim);
            var result:Boolean = match != null && match.length > 0;
            return result ? match: null;
        }

    }
}