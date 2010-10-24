package file
{
        import flash.filesystem.File;
        import flash.utils.Timer;
        import flash.events.TimerEvent;
        import flash.events.EventDispatcher;

        /*
                Todo:

                -Cmonitor changes in multiple attributes
                -add support for monitoring multiple files
        */

        /**
        * Class that monitors files for changes.
        */
        public class FileMonitor extends EventDispatcher
        {
                private var _file:File;
                private var timer:Timer;
                public static const DEFAULT_MONITOR_INTERVAL:Number = 1000;
                private var _interval:Number;
                private var fileExists:Boolean = false;

                private var lastModifiedTime:Number;
                private var lastSize:Number;

                /**
                 *  Constructor
                 *
                 *      @parameter file The File that will be monitored for changes.
                 *
                 *      @param interval How often in milliseconds the file is polled for
                 *      change events. Default value is 1000, minimum value is 1000
                 */
                public function FileMonitor(file:File = null, interval:Number = -1)
                {
                        this.file = file;

                        if(interval != -1)
                        {
                                if(interval < 1000)
                                {
                                        _interval = 1000;
                                }
                                else
                                {
                                        _interval = interval;
                                }
                        }
                        else
                        {
                                _interval = DEFAULT_MONITOR_INTERVAL;
                        }
                }

                /**
                 * File being monitored for changes.
                 *
                 * Setting the property will result in unwatch() being called.
                 */
                public function get file():File
                {
                        return _file;
                }

                public function set file(file:File):void
                {
                        if(timer && timer.running)
                        {
                                unwatch();
                        }

                        _file = file;

                        if(!_file)
                        {
                                fileExists = false;
                                return;
                        }

                        //note : this will throw an error if new File() is passed in.
                        fileExists = _file.exists;
                        if(fileExists)
                        {
                                lastModifiedTime = _file.modificationDate.getTime();
                        }

                }

                /**
                 *      How often the system is polled for Volume change events.
                 */
                public function get interval():Number
                {
                        return _interval;
                }

                /**
                 * Begins monitoring the specified file for changes.
                 *
                 * Broadcasts Event.CHANGE event when the file's modification date has changed.
                 */
                public function watch():void
                {
                        if(!file)
                        {
                                //should we throw an error?
                                return;
                        }

                        if(timer && timer.running)
                        {
                                return;
                        }

                        //check and see if timer is active. if it is, return
                        if(!timer)
                        {
                                timer = new Timer(_interval);
                                timer.addEventListener(TimerEvent.TIMER, onTimerEvent, false, 0, true);
                        }

                        timer.start();
                }

                /**
                 * Stops watching the specified file for changes.
                 */
                public function unwatch():void
                {
                        if(!timer)
                        {
                                return;
                        }

                        timer.stop();
                        timer.removeEventListener(TimerEvent.TIMER, onTimerEvent);
                }

                private function onTimerEvent(e:TimerEvent):void
                {
                        var outEvent:FileMonitorEvent;

                        if(fileExists != _file.exists)
                        {
                                if(_file.exists)
                                {
                                        //file was created
                                        outEvent = new FileMonitorEvent(FileMonitorEvent.CREATE);
                                        lastModifiedTime = _file.modificationDate.getTime();
                                }
                                else
                                {
                                        //file was moved / deleted
                                        outEvent = new FileMonitorEvent(FileMonitorEvent.MOVE);
                                        unwatch();
                                }
                                fileExists = _file.exists;
                        }
                        else
                        {
                                if(!_file.exists)
                                {
                                        return;
                                }

                                var modifiedTime:Number = _file.modificationDate.getTime();
                                var size:Number = _file.size;

                                if(modifiedTime == lastModifiedTime && size == lastSize)
                                {
                                        return;
                                }

                                lastModifiedTime = modifiedTime;
                                lastSize = size;

                                //file modified
                                outEvent = new FileMonitorEvent(FileMonitorEvent.CHANGE);
                        }

                        if(outEvent)
                        {
                                outEvent.file = _file;
                                dispatchEvent(outEvent);
                        }

                }
        }
}
