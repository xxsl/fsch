package styles
{
    import mx.styles.CSSStyleDeclaration;
    import mx.styles.StyleManager;

    public class CSSIconUtil
    {
        public static function getClass(styleName:String, iconRule:String = "icon"):Class
        {
            var declaration:CSSStyleDeclaration = StyleManager.getStyleDeclaration(styleName);
            if(!declaration)
            {
                trace("Style declaration not found: " + styleName);
                return null;
            }
            else
            {
                var prop:* = declaration.getStyle(iconRule);
                if(!prop)
                {
                    trace("Style declaration: " + styleName + ", styles property: " + iconRule + " is null");
                }
                return prop;
            }
        }
    }
}