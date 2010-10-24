package log
{
    import mx.collections.ArrayCollection;
    import mx.controls.treeClasses.DefaultDataDescriptor;

    public class ObjectDataDescriptor extends DefaultDataDescriptor
    {
        public function ObjectDataDescriptor()
        {
            super();
        }

        override public function isBranch(node:Object, model:Object = null):Boolean
        {
            return node.hasOwnProperty("children") && ArrayCollection(node["children"]).length > 0;
        }
    }
}