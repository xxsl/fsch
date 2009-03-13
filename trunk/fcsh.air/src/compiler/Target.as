package compiler {
	
	[Bindable]
    public class Target
    {
        public var name:String;
        public var outputPath:String;
        public var flexConfig:String;
        public var fileName:String;
        
        public static function fromObject(val:Object):Target
        {
        	var newTarget:Target = new Target();
        	newTarget.name = val.name;
        	newTarget.outputPath = val.outputPath;
        	newTarget.flexConfig = val.flexConfig;
        	newTarget.fileName = val.fileName;
        	return newTarget;
        }
        
        public function command():String
        {
        	return "mxmlc -output=" + outputPath + "\\" + fileName + " -load-config+=" + flexConfig;
        }
    }
}