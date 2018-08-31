# CGpFont Class

The **CGpFont** class allows the creation of **Font** objects. The Font object encapsulates the characteristics, such as family, height, size, and style (or combination of styles), of a specific font. A **Font** object is used when drawing strings.

**Inherits from**: CGpBase.
**Imclude file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsImage) | Creates a Font object. |
| [Clone](#Clone) | Copies the contents of the existing Font object into a new Font object. |
| [GetFamily](#GetFamily) | Gets the font family on which this font is based. |
| [GetHeight](#GetHeight) | Gets the line spacing of this font in the current unit of a specified Graphics object. |
| [GetLogFontA](#GetLogFontA) | Uses a LOGFONTA structure to get the attributes of this Font object. |
| [GetLogFontW](#GetLogFontW) | Uses a LOGFONTW structure to get the attributes of this Font object. |
| [GetSize](#GetSize) | Returns the font size (commonly called the em size) of this Font object. |
| [GetStyle](#GetStyle) | Returns the font size (commonly called the em size) of this Font object. |
| [GetUnit](#GetUnit) | Returns the unit of measure of this Font object. |
| [IsAvailable](#IsAvailable) | Determines whether this Font object was created successfully. |

# CGpFontCollection Class

The **CGpFontCollection** class contains methods for enumerating the font families in a collection of fonts. Objects built from this class include the **InstalledFontCollection** and **PrivateFontCollection**.

**Inherits from**: CGpBase.
**Imclude file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [GetFamilies](#GetFamilies) | Gets the font families contained in this font collection. |
| [GetFamilyCount](#GetFamilyCount) | Gets the number of font families contained in this font collection. |

# CGpInstalledFontCollection Class

Extends **CGpFontCollection**. Does not implement new additional methods. 

The **InstalledFontCollection** object represents the fonts installed on the system.

**Inherits from**: CGpFontCollection.
**Imclude file**: CGpFont.inc.

# CGpPrivateFontCollection Class

Extends **CGpFontCollection**. The **PrivateFontCollection** object is a collection for fonts. This object keeps a collection of fonts specifically for an application. The fonts in the collection can include installed fonts as well as fonts that have not been installed on the system.

**Inherits from**: CGpFontCollection.
**Imclude file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [AddFontFile](#AddFontFile) | Adds a font file to this private font collection. |
| [AddMemoryFont](#AddMemoryFont) | Adds a font that is contained in system memory to a Windows GDI+ font collection. |

# CGpFontFamily Class

The **CGpFontFamily** class encapsulates a set of fonts that make up a font family. A font family is a group of fonts that have the same typeface but different styles.

**Inherits from**: CGpBase.
**Imclude file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorFontFamily) | Creates a FontFamily object based on a specified font collection. |
| [Clone](#Clone) | Copies the contents of the existing FontFamily object into a new FontFamily object. |
| [GenericMonospace](#GenericMonospace) | Gets a FontFamily object that a specifies a generic monospace typeface. |
| [GenericSansSerif](#GenericSansSerif) | Gets a FontFamily object that a specifies a generic sans serif typeface. |
| [GenericSerif](#GenericSerif) | Gets a FontFamily object that a specifies a generic serif typeface. |
| [GetCellAscent](#GetCellAscent) | Gets the cell ascent, in design units, of this font family for the specified style or style combination. |
| [GetCellDescent](#GetCellDescent) | Gets the cell descent, in design units, of this font family for the specified style or style combination. |
| [GetEmHeight](#GetEmHeight) | Gets the size (commonly called em size or em height), in design units, of this font family. |
| [GetFamilyName](#GetFamilyName) | Gets the name of this font family. |
| [GetLineSpacing](#GetLineSpacing) | Gets the line spacing, in design units, of this font family for the specified style or style combination. |
| [IsAvailable](#IsAvailable) | Determines whether this Font object was created successfully. |
| [IsStyleAvailable](#IsStyleAvailable) | Determines whether the specified style is available for this font family. |
