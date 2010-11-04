package ui
{
    import flash.events.MouseEvent;

    import mx.controls.CheckBox;
    import mx.core.IFactory;
    import mx.core.WindowedApplication;
    import mx.core.windowClasses.TitleBar;

    import mx.managers.ToolTipManager;

    import styles.Images;

    /**
     *@author aturtsevitch
     *@date   Nov 4, 2010
     *@time   11:52:49 AM
     *@langversion ActionScript 3.0
     */
    public class TitleBarEx extends TitleBar implements IFactory
    {
        public var alwaysOnTop:CheckBox;

        public function TitleBarEx()
        {
            titleIcon = Images.TITLEBAR_ICON;
            super();
        }


        protected override function createChildren():void
        {
            super.createChildren();
            if (!alwaysOnTop)
            {
                alwaysOnTop = new CheckBox();
                alwaysOnTop.styleName = "alwaysOnTop";
                updateAlwaysOnTopTooltip();
                alwaysOnTop.addEventListener(MouseEvent.CLICK, onAlwaysOnTopClick, false, 0, true);
                addChild(alwaysOnTop);
            }
        }

        private function onAlwaysOnTopClick(event:MouseEvent):void
        {
            updateAlwaysOnTopTooltip();
            WindowedApplication(this.parent).alwaysInFront = alwaysOnTop.selected;
        }

        private function updateAlwaysOnTopTooltip():void
        {
            alwaysOnTop.toolTip = "Always on top [" + (alwaysOnTop.selected ? "on" : "off") + "]";
        }

        override protected function placeButtons(align:String, unscaledWidth:Number, unscaledHeight:Number, leftOffset:Number, rightOffset:Number, cornerOffset:Number):void
        {
            super.placeButtons(align, unscaledWidth, unscaledHeight, leftOffset, rightOffset, cornerOffset);

            var pad:Number = getStyle("buttonPadding");
            var edgePad:Number = getStyle("titleBarButtonPadding");

            if (alwaysOnTop)
            {
                alwaysOnTop.setActualSize(alwaysOnTop.measuredWidth, alwaysOnTop.measuredHeight);

                if (align == "right")
                {
                    alwaysOnTop.move(
                            unscaledWidth - (alwaysOnTop.measuredWidth + minimizeButton.measuredWidth +
                                    maximizeButton.measuredWidth + closeButton.measuredWidth +
                                    (3 * pad)) - cornerOffset - edgePad,
                            (unscaledHeight - alwaysOnTop.measuredHeight) / 2);
                }
                else
                {
                    edgePad = Math.max(edgePad, leftOffset);
                    alwaysOnTop.move(
                            edgePad + (pad * 3) +
                                    maximizeButton.measuredWidth + closeButton.measuredWidth + minimizeButton.measuredWidth,
                            (unscaledHeight - alwaysOnTop.measuredHeight) / 2);
                }
            }
        }

        public function newInstance():*
        {
            return new TitleBarEx();
        }
    }
}