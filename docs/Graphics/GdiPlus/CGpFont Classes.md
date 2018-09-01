# CGpFont Class

The **CGpFont** class allows the creation of **Font** objects. The Font object encapsulates the characteristics, such as family, height, size, and style (or combination of styles), of a specific font. A **Font** object is used when drawing strings.

**Inherits from**: CGpBase.<br>
**Include file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsFont) | Creates a Font object. |
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

**Inherits from**: CGpBase.<br>
**Include file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [GetFamilies](#GetFamilies) | Gets the font families contained in this font collection. |
| [GetFamilyCount](#GetFamilyCount) | Gets the number of font families contained in this font collection. |

# CGpInstalledFontCollection Class

Extends **CGpFontCollection**. Does not implement new additional methods. 

The **InstalledFontCollection** object represents the fonts installed on the system.

**Inherits from**: CGpFontCollection.<br>
**Include file**: CGpFont.inc.

# CGpPrivateFontCollection Class

Extends **CGpFontCollection**. The **PrivateFontCollection** object is a collection for fonts. This object keeps a collection of fonts specifically for an application. The fonts in the collection can include installed fonts as well as fonts that have not been installed on the system.

**Inherits from**: CGpFontCollection.<br>
**Include file**: CGpFont.inc.

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

# <a name="ConstructorsFont"></a>Constructors (CGpFont)

Creates a **Font** object based on the Windows Graphics Device Interface (GDI) font object that is currently selected into a specified device context. This constructor is provided for compatibility with GDI.

```
CONSTRUCTOR CGpFont (BYVAL hdc AS HDC)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to a Windows device context that has a font selected. A handle is a number that Windows uses internally to reference an object. |

Creates a **Font** object indirectly from a Windows Graphics Device Interface (GDI) logical font by using a handle to a GDI LOGFONT structure.

```
CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL hFont AS HFONT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to a Windows device context. A handle is a number that Windows uses internally to reference an object. |
| *hFont* | Handle to a logical font. |

Creates a **Font** object directly from a Windows Graphics Device Interface (GDI) logical font. This constructor is provided for compatibility with GDI.

```
CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTA PTR)
CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTW PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to a Windows device context. A handle is a number that Windows uses internally to reference an object. |
| *plogfont* | Pointer to a LOGFONT structure variable that contains attributes of the font. |

Creates a **Font** object based on a **FontFamily** object, a size, a font style, and a unit of measurement.

```
CONSTRUCTOR CGpFont (BYVAL pFamily AS CGpFontFamily PTR, BYVAL emSize AS SINGLE, _
   BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFamily* | Pointer to a **FontFamily** object that specifies information such as the string that identifies the font family and the font family's text metrics measured in design units. |
| *emSize* | The *em* size of the font measured in the units specified in the unit parameter. |
| *nStyle* | Optional. Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout sets the style as a combination of the three styles. The default value is FontStyleRegular. |
| *nUnit* | Optional. Element of the **Unit** enumeration that specifies the unit of measurement for the font size. The default value is UnitPoint. |

Creates a **Font** object based on a font family, a size, a font style, and a unit of measurement.

```
CONSTRUCTOR CGpFont (BYVAL pwszFamilyName AS WSTRING PTR, BYVAL emSize AS SINGLE, _
   BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0, BYVAL pFontCollection AS CGpFontCollection PTR = NULL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFamilyName* | Name of the font family. |
| *emSize* | The *em* size of the font measured in the units specified in the unit parameter. |
| *nStyle* | Optional. Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout sets the style as a combination of the three styles. The default value is FontStyleRegular. |
| *nUnit*  Optional.| Element of the **Unit** enumeration that specifies the unit of measurement for the font size. The default value is UnitPoint. |
| *pFontCollection* | Optional. Pointer to a **FontCollection** object that specifies a user-defined group of fonts. If the value of this parameter is NULL, the system font collection is used. The default value is NULL. |

# <a name="Clone"></a>Clone (CGpFont)

Copies the contents of the existing **Font** object into a new **Font** object.

```
FUNCTION Clone (BYVAL pCloneFont AS CGpFont PTR) AS GpStatus
```

```
| Parameter  | Description |
| ---------- | ----------- |
| *pCloneFont* | Pointer to the **Font** object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Font object, clones it, and then uses the clone to draw text.
' ========================================================================================
SUB Example_CloneFont (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Font object.
   DIM font AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(16) / rxRatio)

   ' // Create a clone of the Font object
   DIM cloneFont AS CGpFont
   font.Clone(@cloneFont)

   ' // Draw Text with cloneFont.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM wszText AS WSTRING * 260 = "This is a cloned Font"
   graphics.DrawString(@wszText, -1, @cloneFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```
