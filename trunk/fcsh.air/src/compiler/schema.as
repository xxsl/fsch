private var flex_config_xsd_xml:XML =
        <xs:schema id="data" elementFormDefault="qualified"
                xmlns="http://www.adobe.com/2006/flex-config"
                targetNamespace="http://www.adobe.com/2006/flex-config"
                xmlns:xs="http://www.w3.org/2001/XMLSchema">
                <xs:element name="flex-config">
                        <xs:annotation>
                                <xs:documentation>Flex compiler configuration</xs:documentation>
                                </xs:annotation>
                        <xs:complexType>
                                <xs:choice maxOccurs="unbounded" minOccurs="0">
                                        <xs:element type="xs:boolean" name="benchmark">
                                                <xs:annotation>
                                                        <xs:documentation>output performance benchmark</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element name="compiler">
                                                <xs:complexType>
                                                        <xs:choice maxOccurs="unbounded" minOccurs="0">
                                                                <xs:element type="xs:boolean" name="accessible">
                                                                        <xs:annotation>
                                                                                <xs:documentation>generate an accessible SWF</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:string" name="actionscript-file-encoding">
                                                                        <xs:annotation>
                                                                                <xs:documentation>specifies actionscript file encoding. If there is no BOM in the AS3 source files, the compiler will use this file encoding.</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="allow-source-path-overlap">
                                                                        <xs:annotation>
                                                                                <xs:documentation>checks if a source-path entry is a subdirectory of another source-path entry. It helps make the package names of MXML components unambiguous. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="as3">
                                                                        <xs:annotation>
                                                                                <xs:documentation>use the ActionScript 3 class based object model for greater performance and better error reporting. In the class based object model most built-in functions are implemented as fixed methods of classes. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:string" name="context-root">
                                                                        <xs:annotation>
                                                                                <xs:documentation>path to replace context.root tokens for service channel endpoints</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="debug">
                                                                        <xs:annotation>
                                                                                <xs:documentation>generates a movie that is suitable for debugging</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="defaults-css-files">
                                                                        <xs:annotation>
                                                                                <xs:documentation>defaults css files (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="filename" maxOccurs="unbounded" minOccurs="0"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:string" name="defaults-css-url">
                                                                        <xs:annotation>
                                                                                <xs:documentation>defines the location of the default style sheet. Setting this option overrides the implicit use of the defaults.css style sheet in the framework.swc file. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="define">
                                                                        <xs:annotation>
                                                                                <xs:documentation>define a global AS3 conditional compilation definition (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="name"/>
                                                                                        <xs:element type="xs:string" name="value"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="es">
                                                                        <xs:annotation>
                                                                                <xs:documentation>use the ECMAScript edition 3 prototype based object model to allow dynamic overriding of prototype properties. In the prototype based object model built-in functions are implemented as dynamic properties of prototype objects. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="external-library-path">
                                                                        <xs:annotation>
                                                                                <xs:documentation>list of SWC files or directories to compile against but to omit from linking</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="path-element" maxOccurs="unbounded" minOccurs="0"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element name="fonts">
                                                                        <xs:complexType>
                                                                                <xs:choice maxOccurs="unbounded" minOccurs="0">
                                                                                        <xs:element type="xs:boolean" name="advanced-anti-aliasing">
                                                                                                <xs:annotation>
                                                                                                        <xs:documentation>enables advanced anti-aliasing for embedded fonts, which provides greater clarity for small fonts.</xs:documentation>
                                                                                                        </xs:annotation>
                                                                                                </xs:element>
                                                                                        <xs:element type="xs:boolean" name="flash-type">
                                                                                                <xs:annotation>
                                                                                                        <xs:documentation>enables FlashType for embedded fonts, which provides greater clarity for small fonts.</xs:documentation>
                                                                                                        </xs:annotation>
                                                                                                </xs:element>
                                                                                        <xs:element name="languages">
                                                                                                <xs:complexType>
                                                                                                        <xs:sequence maxOccurs="unbounded" minOccurs="0">
                                                                                                                <xs:element name="language-range">
                                                                                                                        <xs:annotation>
                                                                                                                                <xs:documentation>a range to restrict the number of font glyphs embedded into the SWF (advanced)</xs:documentation>
                                                                                                                                </xs:annotation>
                                                                                                                        <xs:complexType>
                                                                                                                                <xs:sequence>
                                                                                                                                        <xs:element type="xs:string" name="lang"/>
                                                                                                                                        <xs:element type="xs:string" name="range"/>
                                                                                                                                        </xs:sequence>
                                                                                                                                </xs:complexType>
                                                                                                                        </xs:element>
                                                                                                                </xs:sequence>
                                                                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                                                                        </xs:complexType>
                                                                                                </xs:element>
                                                                                        <xs:element type="xs:string" name="local-fonts-snapshot">
                                                                                                <xs:annotation>
                                                                                                        <xs:documentation>File containing system font data produced by flex2.tools.FontSnapshot. (advanced)</xs:documentation>
                                                                                                        </xs:annotation>
                                                                                                </xs:element>
                                                                                        <xs:element name="managers">
                                                                                                <xs:annotation>
                                                                                                        <xs:documentation>Compiler font manager classes, in policy resolution order (advanced)</xs:documentation>
                                                                                                        </xs:annotation>
                                                                                                <xs:complexType>
                                                                                                        <xs:sequence>
                                                                                                                <xs:element type="xs:string" name="manager-class" maxOccurs="unbounded" minOccurs="0"/>
                                                                                                                </xs:sequence>
                                                                                                        </xs:complexType>
                                                                                                </xs:element>
                                                                                        <xs:element type="xs:unsignedInt" name="max-cached-fonts">
                                                                                                <xs:annotation>
                                                                                                        <xs:documentation>sets the maximum number of fonts to keep in the server cache.  The default value is 20. (advanced)</xs:documentation>
                                                                                                        </xs:annotation>
                                                                                                </xs:element>
                                                                                        <xs:element type="xs:unsignedInt" name="max-glyphs-per-face">
                                                                                                <xs:annotation>
                                                                                                        <xs:documentation>sets the maximum number of character glyph-outlines to keep in the server cache for each font face. The default value is 1000.</xs:documentation>
                                                                                                        </xs:annotation>
                                                                                                </xs:element>
                                                                                        </xs:choice>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="headless-server">
                                                                        <xs:annotation>
                                                                                <xs:documentation>a flag to set when Flex is running on a server without a display (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="include-libraries">
                                                                        <xs:annotation>
                                                                                <xs:documentation>a list of libraries (SWCs) to completely include in the SWF</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="library" maxOccurs="unbounded" minOccurs="0"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="incremental">
                                                                        <xs:annotation>
                                                                                <xs:documentation>enables incremental compilation</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="keep-all-type-selectors">
                                                                        <xs:annotation>
                                                                                <xs:documentation>disables the pruning of unused CSS type selectors (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="keep-as3-metadata">
                                                                        <xs:annotation>
                                                                                <xs:documentation>keep the specified metadata in the SWF (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="name" maxOccurs="unbounded" minOccurs="0"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="keep-generated-actionscript">
                                                                        <xs:annotation>
                                                                                <xs:documentation>save temporary source files generated during MXML compilation (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="library-path">
                                                                        <xs:annotation>
                                                                                <xs:documentation>list of SWC files or directories that contain SWC files</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="path-element" maxOccurs="unbounded" minOccurs="0"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element name="locale">
                                                                        <xs:annotation>
                                                                                <xs:documentation>specifies the locale for internationalization</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="locale-element" maxOccurs="unbounded" minOccurs="0"/>
                                                                                        </xs:sequence>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element name="mxml">
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="compatibility-version">
                                                                                                <xs:annotation>
                                                                                                        <xs:documentation>specifies a compatibility version. e.g. 2.0.1</xs:documentation>
                                                                                                        </xs:annotation>
                                                                                                </xs:element>
                                                                                        </xs:sequence>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element name="namespaces">
                                                                        <xs:complexType>
                                                                                <xs:sequence maxOccurs="unbounded" minOccurs="0">
                                                                                        <xs:element name="namespace">
                                                                                                <xs:annotation>
                                                                                                        <xs:documentation>Specify a URI to associate with a manifest of components for use as MXML elements</xs:documentation>
                                                                                                        </xs:annotation>
                                                                                                <xs:complexType>
                                                                                                        <xs:sequence>
                                                                                                                <xs:element type="xs:string" name="uri"/>
                                                                                                                <xs:element type="xs:string" name="manifest"/>
                                                                                                                </xs:sequence>
                                                                                                        </xs:complexType>
                                                                                                </xs:element>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="optimize">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Enable post-link SWF optimization</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:string" name="services">
                                                                        <xs:annotation>
                                                                                <xs:documentation>path to Flex Data Services configuration file</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="show-actionscript-warnings">
                                                                        <xs:annotation>
                                                                                <xs:documentation>runs the AS3 compiler in a mode that detects legal but potentially incorrect code</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="show-binding-warnings">
                                                                        <xs:annotation>
                                                                                <xs:documentation>toggle whether warnings generated from data binding code are displayed</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="show-shadowed-device-font-warnings">
                                                                        <xs:annotation>
                                                                                <xs:documentation>toggles whether warnings are displayed when an embedded font name shadows a device font name</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="show-unused-type-selector-warnings">
                                                                        <xs:annotation>
                                                                                <xs:documentation>toggle whether warnings generated from unused CSS type selectors are displayed</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="source-path">
                                                                        <xs:annotation>
                                                                                <xs:documentation>list of path elements that form the roots of ActionScript class hierarchies</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="path-element" maxOccurs="unbounded" minOccurs="0"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="strict">
                                                                        <xs:annotation>
                                                                                <xs:documentation>runs the AS3 compiler in strict error checking mode.</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="theme">
                                                                        <xs:annotation>
                                                                                <xs:documentation>list of CSS or SWC files to apply as a theme</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="filename" maxOccurs="unbounded" minOccurs="0"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="use-resource-bundle-metadata">
                                                                        <xs:annotation>
                                                                                <xs:documentation>determines whether resources bundles are included in the application.</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="verbose-stacktraces">
                                                                        <xs:annotation>
                                                                                <xs:documentation>save callstack information to the SWF for debugging (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-array-tostring-changes">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Array.toString() format has changed. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-assignment-within-conditional">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Assignment within conditional. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-array-cast">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Possibly invalid Array cast operation. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-bool-assignment">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Non-Boolean value used where a Boolean value was expected. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-date-cast">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Invalid Date cast operation. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-es3-type-method">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Unknown method. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-es3-type-prop">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Unknown property. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-nan-comparison">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Illogical comparison with NaN. Any comparison operation involving NaN will evaluate to false because NaN != NaN. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-null-assignment">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Impossible assignment to null. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-null-comparison">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Illogical comparison with null. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-bad-undefined-comparison">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Illogical comparison with undefined.  Only untyped variables (or variables of type *) can be undefined. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-boolean-constructor-with-no-args">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Boolean() with no arguments returns false in ActionScript 3.0. Boolean() returned undefined in ActionScript 2.0. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-changes-in-resolve">
                                                                        <xs:annotation>
                                                                                <xs:documentation>__resolve is no longer supported. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-class-is-sealed">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Class is sealed.  It cannot have members added to it dynamically. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-const-not-initialized">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Constant not initialized. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-constructor-returns-value">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Function used in new expression returns a value.  Result will be what the function returns, rather than a new instance of that function. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-deprecated-event-handler-error">
                                                                        <xs:annotation>
                                                                                <xs:documentation>EventHandler was not added as a listener. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-deprecated-function-error">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Unsupported ActionScript 2.0 function. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-deprecated-property-error">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Unsupported ActionScript 2.0 property. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-duplicate-argument-names">
                                                                        <xs:annotation>
                                                                                <xs:documentation>More than one argument by the same name. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-duplicate-variable-def">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Duplicate variable definition  (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-for-var-in-changes">
                                                                        <xs:annotation>
                                                                                <xs:documentation>ActionScript 3.0 iterates over an object's properties within a "for x in target" statement in random order. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-import-hides-class">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Importing a package by the same name as the current class will hide that class identifier in this scope. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-instance-of-changes">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Use of the instanceof operator. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-internal-error">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Internal error in compiler. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-level-not-supported">
                                                                        <xs:annotation>
                                                                                <xs:documentation>_level is no longer supported. For more information, see the flash.display package. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-missing-namespace-decl">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Missing namespace declaration (e.g. variable is not defined to be public, private, etc.). (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-negative-uint-literal">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Negative value will become a large positive value when assigned to a uint data type. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-no-constructor">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Missing constructor. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-no-explicit-super-call-in-constructor">
                                                                        <xs:annotation>
                                                                                <xs:documentation>The super() statement was not called within the constructor. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-no-type-decl">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Missing type declaration. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-number-from-string-changes">
                                                                        <xs:annotation>
                                                                                <xs:documentation>In ActionScript 3.0, white space is ignored and '' returns 0. Number() returns NaN in ActionScript 2.0 when the parameter is '' or contains white space. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-scoping-change-in-this">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Change in scoping for the this keyword.  Class methods extracted from an instance of a class will always resolve this back to that instance.  In ActionScript 2.0 this is looked up dynamically based on where the method is invoked from. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-slow-text-field-addition">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Inefficient use of += on a TextField. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-unlikely-function-value">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Possible missing parentheses. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:boolean" name="warn-xml-class-has-changed">
                                                                        <xs:annotation>
                                                                                <xs:documentation>Possible usage of the ActionScript 2.0 XML class. (advanced)</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                </xs:choice>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element type="xs:boolean" name="compute-digest">
                                                <xs:annotation>
                                                        <xs:documentation>writes a digest to the catalog.xml of a library. This is required when the library will be used in the -runtime-shared-libraries-path option.</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:string" name="debug-password">
                                                <xs:annotation>
                                                        <xs:documentation>the password to include in debuggable SWFs (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:unsignedInt" name="default-background-color">
                                                <xs:annotation>
                                                        <xs:documentation>default background color (may be overridden by the application code) (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:unsignedInt" name="default-frame-rate">
                                                <xs:annotation>
                                                        <xs:documentation>default frame rate to be used in the SWF. (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element name="default-script-limits">
                                                <xs:annotation>
                                                        <xs:documentation>default script execution limits (may be overridden by root attributes) (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:unsignedInt" name="max-recursion-depth"/>
                                                                <xs:element type="xs:unsignedInt" name="max-execution-time"/>
                                                                </xs:sequence>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="default-size">
                                                <xs:annotation>
                                                        <xs:documentation>default application size (may be overridden by root attributes in the application) (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:unsignedInt" name="width"/>
                                                                <xs:element type="xs:unsignedInt" name="height"/>
                                                                </xs:sequence>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element type="xs:boolean" name="directory">
                                                <xs:annotation>
                                                        <xs:documentation>output the library as an open directory instead of a SWC file</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:string" name="dump-config">
                                                <xs:annotation>
                                                        <xs:documentation>write a file containing all currently set configuration values in a format suitable for use as a flex config file (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element name="externs">
                                                <xs:annotation>
                                                        <xs:documentation>a list of symbols to omit from linking when building a SWF (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence maxOccurs="unbounded" minOccurs="0">
                                                                <xs:element type="xs:string" name="symbol"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="frames">
                                                <xs:annotation>
                                                        <xs:documentation>A SWF frame label with a sequence of classnames that will be linked onto the frame. (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence maxOccurs="unbounded" minOccurs="0">
                                                                <xs:element name="frame">
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="label"/>
                                                                                        <xs:element type="xs:string" name="classname"/>
                                                                                        </xs:sequence>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="include-classes">
                                                <xs:annotation>
                                                        <xs:documentation>a list of classes to include in the output SWC</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="class" maxOccurs="unbounded" minOccurs="0"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="include-file">
                                                <xs:annotation>
                                                        <xs:documentation>a list of named files to include in the output SWC</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="name"/>
                                                                <xs:element type="xs:string" name="path"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element type="xs:boolean" name="include-lookup-only">
                                                <xs:annotation>
                                                        <xs:documentation>if true, manifest entries with lookupOnly=true are included in SWC catalog. Default is false. (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element name="include-namespaces">
                                                <xs:annotation>
                                                        <xs:documentation>all classes in the listed namespaces are included in the output SWC</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="uri" maxOccurs="unbounded" minOccurs="0"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="include-resource-bundles">
                                                <xs:annotation>
                                                        <xs:documentation>a list of resource bundles to include in the output SWC</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="bundle" maxOccurs="unbounded" minOccurs="0"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="include-sources">
                                                <xs:annotation>
                                                        <xs:documentation>a list of directories and source files to include in the output SWC</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="path-element" maxOccurs="unbounded" minOccurs="0"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="include-stylesheet">
                                                <xs:annotation>
                                                        <xs:documentation>a list of named stylesheet resources to include in the output SWC</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="name"/>
                                                                <xs:element type="xs:string" name="path"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="includes">
                                                <xs:annotation>
                                                        <xs:documentation>a list of symbols to always link in when building a SWF (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="symbol" maxOccurs="unbounded" minOccurs="0"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="licenses">
                                                <xs:annotation>
                                                        <xs:documentation>specifies a product and a serial number.</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence maxOccurs="unbounded" minOccurs="0">
                                                                <xs:element name="license">
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="product"/>
                                                                                        <xs:element type="xs:string" name="serial-number"/>
                                                                                        </xs:sequence>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element type="xs:string" name="link-report">
                                                <xs:annotation>
                                                        <xs:documentation>Output a XML-formatted report of all definitions linked into the application. (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element name="load-config">
                                                <xs:annotation>
                                                        <xs:documentation>load a file containing configuration options</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:simpleContent>
                                                                <xs:extension base="xs:string">
                                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                                        </xs:extension>
                                                                </xs:simpleContent>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="load-externs">
                                                <xs:annotation>
                                                        <xs:documentation>an XML file containing &lt;def>, &lt;pre>, and &lt;ext> symbols to omit from linking when building a SWF (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:simpleContent>
                                                                <xs:extension base="xs:string">
                                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                                        </xs:extension>
                                                                </xs:simpleContent>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="metadata">
                                                <xs:complexType>
                                                        <xs:choice maxOccurs="unbounded" minOccurs="0">
                                                                <xs:element name="contributor">
                                                                        <xs:annotation>
                                                                                <xs:documentation>A contributor's name to store in the SWF metadata</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:simpleContent>
                                                                                        <xs:extension base="xs:string">
                                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                                </xs:extension>
                                                                                        </xs:simpleContent>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element name="creator">
                                                                        <xs:annotation>
                                                                                <xs:documentation>A creator's name to store in the SWF metadata</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:simpleContent>
                                                                                        <xs:extension base="xs:string">
                                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                                </xs:extension>
                                                                                        </xs:simpleContent>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:string" name="date">
                                                                        <xs:annotation>
                                                                                <xs:documentation>The creation date to store in the SWF metadata</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element type="xs:string" name="description">
                                                                        <xs:annotation>
                                                                                <xs:documentation>The default description to store in the SWF metadata</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                <xs:element name="language">
                                                                        <xs:annotation>
                                                                                <xs:documentation>The language to store in the SWF metadata (i.e. EN, FR)</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:simpleContent>
                                                                                        <xs:extension base="xs:string">
                                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                                </xs:extension>
                                                                                        </xs:simpleContent>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element name="localized-description">
                                                                        <xs:annotation>
                                                                                <xs:documentation>A localized RDF/XMP description to store in the SWF metadata</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="text"/>
                                                                                        <xs:element type="xs:string" name="lang"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element name="localized-title">
                                                                        <xs:annotation>
                                                                                <xs:documentation>A localized RDF/XMP title to store in the SWF metadata</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:sequence>
                                                                                        <xs:element type="xs:string" name="title"/>
                                                                                        <xs:element type="xs:string" name="lang"/>
                                                                                        </xs:sequence>
                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element name="publisher">
                                                                        <xs:annotation>
                                                                                <xs:documentation>A publisher's name to store in the SWF metadata</xs:documentation>
                                                                                </xs:annotation>
                                                                        <xs:complexType>
                                                                                <xs:simpleContent>
                                                                                        <xs:extension base="xs:string">
                                                                                                <xs:attribute type="xs:boolean" name="append"/>
                                                                                                </xs:extension>
                                                                                        </xs:simpleContent>
                                                                                </xs:complexType>
                                                                        </xs:element>
                                                                <xs:element type="xs:string" name="title">
                                                                        <xs:annotation>
                                                                                <xs:documentation>The default title to store in the SWF metadata</xs:documentation>
                                                                                </xs:annotation>
                                                                        </xs:element>
                                                                </xs:choice>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element type="xs:string" name="output">
                                                <xs:annotation>
                                                        <xs:documentation>the filename of the SWF movie to create</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:string" name="raw-metadata">
                                                <xs:annotation>
                                                        <xs:documentation>XML text to store in the SWF metadata (overrides metadata.* configuration) (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:string" name="resource-bundle-list">
                                                <xs:annotation>
                                                        <xs:documentation>prints a list of resource bundles to a file for input to the compccompiler to create a resource bundle SWC file. (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element name="runtime-shared-libraries">
                                                <xs:annotation>
                                                        <xs:documentation>a list of runtime shared library URLs to be loaded before the application starts</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="url" maxOccurs="unbounded" minOccurs="0"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element name="runtime-shared-library-path">
                                                <xs:complexType>
                                                        <xs:choice maxOccurs="unbounded" minOccurs="0">
                                                                <xs:element type="xs:string" name="path-element"/>
                                                                <xs:element type="xs:string" name="rsl-url"/>
                                                                <xs:element type="xs:string" name="policy-file-url"/>
                                                                </xs:choice>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        <xs:element type="xs:boolean" name="static-link-runtime-shared-libraries">
                                                <xs:annotation>
                                                        <xs:documentation>statically link the libraries specified by the -runtime-shared-libraries-path option.</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:string" name="target-player">
                                                <xs:annotation>
                                                        <xs:documentation>specifies the version of the player the application is targeting. Features requiring a later version will not be compiled into the application. The minimum value supported is "9.0.0".</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:boolean" name="use-network">
                                                <xs:annotation>
                                                        <xs:documentation>toggle whether the SWF is flagged for access to network resources</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:boolean" name="verify-digests">
                                                <xs:annotation>
                                                        <xs:documentation>verifies the libraries loaded at runtime are the correct ones. (advanced)</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:boolean" name="version">
                                                <xs:annotation>
                                                        <xs:documentation>display the build version of the program</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>
                                        <xs:element type="xs:boolean" name="warnings">
                                                <xs:annotation>
                                                        <xs:documentation>toggle the display of warnings</xs:documentation>
                                                        </xs:annotation>
                                                </xs:element>

                                        <xs:element name="file-specs">
                                                <xs:annotation>
                                                        <xs:documentation>a list of directories and source files to include in the output SWC</xs:documentation>
                                                        </xs:annotation>
                                                <xs:complexType>
                                                        <xs:sequence>
                                                                <xs:element type="xs:string" name="path-element" maxOccurs="unbounded" minOccurs="0"/>
                                                                </xs:sequence>
                                                        <xs:attribute type="xs:boolean" name="append"/>
                                                        </xs:complexType>
                                                </xs:element>
                                        </xs:choice>
                                </xs:complexType>
                        </xs:element>
                </xs:schema>;