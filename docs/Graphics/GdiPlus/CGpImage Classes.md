# CGpImage Class

The **CGpImage** class provides methods for loading and saving raster images (bitmaps) and vector images (metafiles). An Image object encapsulates a bitmap or a metafile and stores attributes that you can retrieve by calling various Get methods. You can construct **Image** objects from a variety of file types including BMP, ICON, GIF, JPEG, Exif, PNG, TIFF, WMF, and EMF.

**Inherits from**: CGpBase.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsImage) | Create an **Image** object based on a file or stream. |
| [Clone](#Clone) | Copies the contents of the existing **Image** object into a new Image object. |
| [FindFirstItem](#FindFirstItem) | Retrieves the description and the data size of the first metadata item in this Image object. |
| [FindNextItem](#FindNextItem) | Retrieves the description and the data size of the next metadata item in this Image object. |
| [GetAllPropertyItems](#GetAllPropertyItems) | Gets all the property items (metadata) stored in this Image object. |
| [GetBounds](#GetBounds) | Gets the bounding rectangle for this image. |
| [GetEncoderParameterList](#GetEncoderParameterList) | Gets a list of the parameters supported by a specified image encoder. |
| [GetEncoderParameterListSize](#GetEncoderParameterListSize) | Gets the size, in bytes, of the parameter list for a specified image encoder. |
| [GetFlags](#GetFlags) | Gets a set of flags that indicate certain attributes of this Image object. |
| [GetFrameCount](#GetFrameCount) | Gets the number of frames in a specified dimension of this Image object. |
| [GetFrameDimensionsCount](#GetFrameDimensionsCount) | Gets the number of frame dimensions in this Image object. |
| [GetFrameDimensionsList](#GetFrameDimensionsList) | Gets the identifiers for the frame dimensions of this Image object. |
| [GetHeight](#GetHeight) | Gets the image height, in pixels, of this image. |
| [GetHorizontalResolution](#GetHorizontalResolution) | Gets the horizontal resolution, in dots per inch, of this image. |
| [GetItemData](#GetItemData) | Gets one piece of metadata from this Image object. |
| [GetPalette](#GetPalette) | Gets the **ColorPalette** of this Image object. |
| [GetPaletteSize](#GetPaletteSize) | Gets the size, in bytes, of the color palette of this Image object. |
| [GetPhysicalDimension](#GetPhysicalDimension) | Gets the width and height of this image. |
| [GetPixelFormat](#GetPixelFormat) | Gets the pixel format of this Image object. |
| [GetPropertyCount](#GetPropertyCount) | Gets the number of properties (pieces of metadata) stored in this Image object. |
| [GetPropertyIdList](#GetPropertyIdList) | Gets a list of the property identifiers used in the metadata of this Image object. |
| [GetPropertyItem](#GetPropertyItem) | Gets a specified property item (piece of metadata) from this Image object. |
| [GetPropertyItemSize](#GetPropertyItemSize) | Gets the size, in bytes, of a specified property item of this Image object. |
| [GetPropertySize](#GetPropertySize) | Gets the total size, in bytes, of all the property items stored in this Image object. |
| [GetRawFormat](#GetRawFormat) | Gets a globally unique identifier ( GUID) that identifies the format of this Image object. |
| [GetThumbnailImage](#GetThumbnailImage) | Gets a thumbnail image from this Image object. |
| [GetType](#GetType) | Gets the type (bitmap or metafile) of this Image object. |
| [GetVerticalResolution](#GetVerticalResolution) | Gets the vertical resolution, in dots per inch, of this image. |
| [GetWidth](#GetWidth) | Gets the width, in pixels, of this image. |
| [RemovePropertyItem](#RemovePropertyItem) | Removes a property item (piece of metadata) from this Image object. |
| [RotateFlip](#RotateFlip) | Rotates and flips this image. |
| [Save](#Save) | Saves this image to a file. |
| [SaveAdd](#SaveAdd) | Adds a frame to a file or stream specified in a previous call to the Save method. |
| [SelectActiveFrame](#SelectActiveFrame) | Selects the frame in this Image object specified by a dimension and an index. |
| [SetPalette](#SetPalette) | Sets the color palette of this Image object. |
| [SetPropertyItem](#SetPropertyItem) | Sets a property item (piece of metadata) for this Image object. |

# CGpBitmap Class

Extends the **CGpImage** class. The **Bitmap** object expands on the capabilities of the **Image** object by providing additional methods for creating and manipulating raster images.

**Inherits from**: CGpImage.
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsBitmap) | Create a Bitmap object based on an icon or resource file. |
| [Clone](#Clone) | Creates a new Bitmap object by copying a portion of this bitmap. |
| [ConvertFormat](#ConvertFormat) | Converts a bitmap to a specified pixel format. |
| [GetHBITMAP](#GetHBITMAP) | Creates a Windows Graphics Device Interface (GDI) bitmap from this Bitmap object. |
| [GetHICON](#GetHICON) | Creates an icon from this Bitmap object. |
| [GetHistogram](#GetHistogram) | Returns one or more histograms for specified color channels of this Bitmap object. |
| [GetHistogramSize](#GetHistogramSize) | Returns the number of elements (in an array of DWORDs) that you must allocate before you call the GetHistogram method of a Bitmap object. |
| [GetPixel](#GetPixel) | Gets the color of a specified pixel in this bitmap. |
| [InitializePalette](#InitializePalette) | Initializes a standard, optimal, or custom color palette. |
| [LockBits](#LockBits) | Locks a rectangular portion of this bitmap and provides a temporary buffer that you can use to read or write pixel data in a specified format. |
| [SetPixel](#SetPixel) | Sets the color of a specified pixel in this bitmap. |
| [SetResolution](#SetResolution) | Sets the resolution of this Bitmap object. |
| [UnlockBits](#UnlockBits) | Unlocks a portion of this bitmap that was previously locked by a call to LockBits. |

# CGpMetafile Class

Extends the **CGpImage** class. The **Metafile** object defines a graphic metafile. A metafile contains records that describe a sequence of graphics API calls. Metafiles can be recorded (constructed) and played back (displayed).

**Inherits from**: CGpImage.
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsMetafile) | Creates a Windows GDI+ Metafile. |
| [EmfToWmfBits](#EmfToWmfBits) | Converts an enhanced-format metafile to a Windows Metafile Format (WMF) metafile and stores the converted records in a specified buffer. |
| [GetDownLevelRasterizationLimit](#GetDownLevelRasterizationLimit) | Gets the rasterization limit currently set for this metafile. |
| [GetHENHMETAFILE](#GetHENHMETAFILE) | Gets a Windows handle to an Enhanced Metafile (EMF) file. |
| [GetMetafileHeader](#GetMetafileHeader) | Gets the metafile header of this metafile. |
| [PlayRecord](#PlayRecord) | Plays a metafile record. |
| [SetDownLevelRasterizationLimit](#SetDownLevelRasterizationLimit) | Sets the resolution for certain brush bitmaps that are stored in this metafile. |

# CGpCachedBitmap Class

Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. The cached bitmap takes the pixel data from the **Bitmap** object and stores it in a format that is optimized for the display device associated with the **Graphics** object.

**Inherits from**: CGpBase.
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorCachedBitmap) | Creates a CachedBitmap object based on a Bitmap object and a Graphics object. |

# <a name="ConstructorsImage"></a>Constructors (CGpImage)

Creates an **Image** object based on a file.

```
CONSTRUCTOR CGpImage (BYVAL pwszFileName AS WSTRING PTR,  BYVAL useicm AS BOOLEAN = FALSE)
```

Create a **Image** object based on an **IStream** interface.

```
CONSTRUCTOR CGpImage (BYVAL pStream AS IStream PTR, BYVAL useicm AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Pointer to a null-terminated string that specifies the path name of the image file. The graphics file formats supported by GDI+ are BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF. |
| *pStream* | Pointer to an **IStream** interface. |
| *useicm* | Optional. Boolean value that specifies whether the new Image object applies color correction according to color management information that is embedded in the image file. Embedded information can include International Color Consortium (ICC) profiles, gamma values, and chromaticity information. TRUE specifies that color correction is enabled, and FALSE specifies that color correction is not enabled. The default value is FALSE. |

# <a name="Clone"></a>Clone (CGpImage)

Creates an **Image** object based on a file.

```
FUNCTION Clone (BYVAL pCloneImage AS CGpImage PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneImage* | Pointer to the **Image** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a JPEG file. The code creates a
' second Image object by cloning the first. Then the code calls the DrawImage method twice
' to draw the two images.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create an Image object, and then clone it.
   DIM image1 AS CGpImage = "climber.jpg"
   DIM pImage2 AS CGpImage
   image1.Clone(@pImage2)

   ' // Draw the original image and the cloned image.
   graphics.DrawImage(@image1, 20 * rxRatio, 20 * ryRatio)
   graphics.DrawImage(@pImage2, 230 * rxRatio, 20 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="FindFirstItem"></a>FindFirstItem (CGpImage)

Retrieves the description and the data size of the first metadata item in this Image object.

```
FUNCTION FindFirstItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | In, Out. Pointer to an **ImageItemData** structure. On input, the **Desc** member points to a buffer (allocated by the caller) large enough to hold the metadata description (1 byte for JPEG, 4 bytes for PNG, 11 bytes for GIF), and the **DescSize** member specifies the size (1, 4, or 6) of the buffer pointed to by **Desc**. On output, the buffer pointed to by **Desc** receives the metadata description, and the **DataSize** member receives the size, in bytes, of the metadata itself. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Use **FindFirstItem** along with **FindNextItem** to enumerate the metadata items, including custom metadata, stored in an image. **FindFirstItem** and **FindNextItem** do not enumerate the metadata items stored by the **SetPropertyItem** method.

# <a name="FindNextItem"></a>FindNextItem (CGpImage)

The **FindNextItem** method is used along with the **FindFirstItem** method to enumerate the metadata items stored in this **Image** object. The **FindNextItem** method retrieves the description and the data size of the next metadata item in this **Image** object. 

```
FUNCTION FindNextItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | In, Out. Pointer to an **ImageItemData** structure. On input, the **Desc** member points to a buffer (allocated by the caller) large enough to hold the metadata description (1 byte for JPEG, 4 bytes for PNG, 11 bytes for GIF), and the **DescSize** member specifies the size (1, 4, or 6) of the buffer pointed to by **Desc**. On output, the buffer pointed to by **Desc** receives the metadata description, and the **DataSize** member receives the size, in bytes, of the metadata itself. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

**FindFirstItem** and **FindNextItem** do not enumerate the metadata items stored by the **SetPropertyItem** method.

# <a name="GetAllPropertyItems"></a>GetAllPropertyItems (CGpImage)

Gets all the property items (metadata) stored in this **Image** object.

```
FUNCTION GetAllPropertyItems (BYVAL totalBufferSize AS UINT, BYVAL numProperties AS UINT, _
   BYVAL allItems AS PropertyItem PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *totalBufferSize* | Integer that specifies the size, in bytes, of the allItems buffer. Call the **GetPropertySize** method to obtain the required size.  |
| *numProperties* | Integer that specifies the number of properties in the image. Call the **GetPropertySize** method to obtain this number. |
| *allItems* | Pointer to an array of **PropertyItem** structures that receives the property items. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Some image files contain metadata that you can read to determine features of the image. For example, a digital photograph might contain metadata that you can read to determine the make and model of the camera used to capture the image.

GDI+ stores an individual piece of metadata in a PropertyItem object. The **GetAllPropertyItems** method returns an array of PropertyItem objects. Before you call **GetAllPropertyItems**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertySize** method of an **Image** object to get the size, in bytes, of the required buffer. The **GetPropertySize** method also gives you the number of properties (pieces of metadata) in the image.

Several enumerations and constants related to image metadata are defined in Gdiplusimaging.inc.

# <a name="GetBounds"></a>GetBounds (CGpImage)

Gets the bounding rectangle for this image.

```
FUNCTION GetBounds (BYVAL srcRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *srcRect* | Pointer to a **GpRectF** object that receives the bounding rectangle. |
| *srcUnit* | Pointer to a variable that receives an element of the **GpUnit** enumeration that indicates the unit of measure for the bounding rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The bounding rectangle for a metafile does not necessarily have (0, 0) as its upper-left corner. The coordinates of the upper-left corner can be negative or positive, depending on the drawing commands that were issued during the recording of the metafile. For example, suppose a metafile consists of a single ellipse that was recorded with the following statement:

```
DrawEllipse(pen, 200, 100, 80, 40)
```

Then the bounding rectangle for the metafile will enclose that one ellipse. The upper-left corner of the bounding rectangle will not be (0, 0); rather, it will be a point near (200, 100).

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a metafile and then draws the image.
' Next, the code calls the Image.GetBounds method to get the bounding rectangle for the image
' and redraws the a 75 per cent of the image.
' ========================================================================================
SUB Example_GetBounds (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pImage AS CGpImage = "climber.emf"
   graphics.DrawImage(@pImage, 0, 0)

   ' // Get the bounding rectangle for the image (metafile).
   DIM boundsRect AS RectF
   DIM nUnit AS GpUnit
   pImage.GetBounds(@boundsRect, @nUnit)

   ' // Draw 75 percent of the image.
   graphics.DrawImage(@pImage, 230.0, 0.0, boundsRect.X, boundsRect.Y, 0.75 * boundsRect.Width, boundsRect.Height, UnitPixel)

END SUB
' ========================================================================================
```

# <a name="GetEncoderParameterList"></a>GetEncoderParameterList (CGpImage)

Gets a list of the parameters supported by a specified image encoder.

```
FUNCTION GetEncoderParameterList (BYVAL clsidEncoder AS GUID PTR, BYVAL uSize AS UINT, _
   BYVAL buffer AS EncoderParameters PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *clsidEncoder* | Pointer to a **CLSID** that specifies the encoder. |
| *uSize* | Integer that specifies the size, in bytes, of the buffer array. Call the **GetEncoderParameterListSize** method to obtain the required size. |
| *buffer* | Pointer to an **EncoderParameters** object that receives the list of supported parameters. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The **GetEncoderParameterList** method returns an array of **EncoderParameter** objects. Before you call **GetEncoderParameterList**, you must allocate a buffer large enough to receive that array, which is part of an **EncoderParameters** object. You can call the **GetEncoderParameterListSize** method to get the size, in bytes, of the required **EncoderParameters** object.

# <a name="GetEncoderParameterListSize"></a>GetEncoderParameterListSize (CGpImage)

Gets the size, in bytes, of the parameter list for a specified image encoder.

```
FUNCTION GetEncoderParameterListSize (BYVAL clsidEncoder AS GUID PTR) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *clsidEncoder* | Pointer to a **CLSID** that specifies the encoder. |

# <a name="GetFlags"></a>GetFlags (CGpImage)

Gets a set of flags that indicate certain attributes of this **Image** object.

```
FUNCTION GetFlags () AS UINT
```

#### Return value

This method returns a value that holds a set of single-bit flags. The flags are defined in the **ImageFlags** enumeration.

# <a name="GetFrameCount"></a>GetFrameCount (CGpImage)

Gets the number of frames in a specified dimension of this **Image** object.

```
FUNCTION GetFrameCount (BYVAL dimensionID AS GUID PTR) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dimensionID* | Pointer to a GUID that specifies the dimension. GUIDs that identify various dimensions are defined in Gdiplusimaging.inc. |

# <a name="GetFrameDimensionsCount"></a>GetFrameDimensionsCount (CGpImage)

Gets the number of frame dimensions in this **Image** object.

```
FUNCTION GetFrameDimensionsCount () AS UINT
```

#### Remarks

This method returns information about multiple-frame images, which come in two styles: multiple page and multiple resolution.

A multiple-page image is an image that contains more than one image. Each page contains a single image (or frame). These pages (or images, or frames) are typically displayed in succession to produce an animated sequence, such as in an animated GIF file.

A multiple-resolution image is an image that contains more than one copy of an image at different resolutions.

Windows GDI+ can support an arbitrary number of pages (or images, or frames), as well as an arbitrary number of resolutions.

# <a name="GetFrameDimensionsList"></a>GetFrameDimensionsList (CGpImage)

Gets the identifiers for the frame dimensions of this Image object.

```
FUNCTION GetFrameDimensionsList (BYVAL dimensionIDs AS GUID PTR, BYVAL count AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dimensionIDs* | Pointer to an array that receives the identifiers. GUIDs that identify various dimensions are defined in Gdiplusimaging.inc. |
| *count* | Integer that specifies the number of elements in the dimensionIDs array. Call the **GetFrameDimensionsCount** method to determine this number. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

This method returns information about multiple-frame images, which come in two styles: multiple page and multiple resolution.

A multiple-page image is an image that contains more than one image. Each page contains a single image (or frame). These pages (or images, or frames) are typically displayed in succession to produce an animated sequence, such as in an animated GIF file.

A multiple-resolution image is an image that contains more than one copy of an image at different resolutions.

Windows GDI+ can support an arbitrary number of pages (or images, or frames), as well as an arbitrary number of resolutions.

# <a name="GetHeight"></a>GetHeight (CGpImage)

Gets the image height, in pixels, of this image.

```
FUNCTION GetHeight () AS UINT
```

# <a name="GetHorizontalResolution"></a>GetHorizontalResolution (CGpImage)

Gets the horizontal resolution, in dots per inch, of this image.

```
FUNCTION GetHorizontalResolution () AS SINGLE
```

# <a name="GetItemData"></a>GetItemData (CGpImage)

Gets one piece of metadata from this **Image** object.

```
FUNCTION GetItemData (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | Pointer to an **ImageItemData** object that specifies the item to be retrieved. The **Data** member of the **ImageItemData** object points to a buffer that receives the custom metadata. If the **Data** member is set to NULL, this method returns the size of the required buffer in the **DataSize** member of the **ImageItemData** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPalette"></a>GetPalette (CGpImage)

Gets the **ColorPalette** of this Image object.

```
FUNCTION GetPalette (BYVAL pal AS ColorPalette PTR, BYVAL nSize AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pal* | Pointer to a **ColorPalette** structure that receives the palette. |
| *nSize* | Integer that specifies the size, in bytes, of the palette. Call the **GetPaletteSize** method to determine the size. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPaletteSize"></a>GetPaletteSize (CGpImage)

Gets the size, in bytes, of the color palette of this **Image** object.

```
FUNCTION GetPaletteSize () AS INT_
```

# <a name="GetPhysicalDimension"></a>GetPhysicalDimension (CGpImage)

Gets the width and height of this image.

```
FUNCTION GetPhysicalDimension (BYVAL psize AS SizeF PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *psize* | Pointer to a **GpSizeF** object that receives the width and height of this image. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPixelFormat"></a>GetPixelFormat (CGpImage)

Gets the pixel format of this **Image** object.

```
FUNCTION GetPixelFormat () AS PixelFormat
```

#### Return value

This method returns an integer that indicates the pixel format of this **Image** object. The **PixelFormat** data type and constants that represent various pixel formats are defined in Gdipluspixelformats.inc.

# <a name="GetPropertyCount"></a>GetPropertyCount (CGpImage)

Gets the number of properties (pieces of metadata) stored in this **Image** object.

```
FUNCTION GetPropertyCount () AS UINT
```

# <a name="GetPropertyIdList"></a>GetPropertyIdList (CGpImage)

Gets a list of the property identifiers used in the metadata of this **Image** object.

```
FUNCTION GetPropertyIdList (BYVAL numOfProperty AS UINT, BYVAL list AS PROPID PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *numOfProperty* | Integer that specifies the number of elements in the list array. Call the **GetPropertyCount** method to determine this number. |
| *list* | Pointer to an array that receives the property identifiers. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The **GetPropertyIdList** method returns an array of **PROPIDs**. Before you call **GetPropertyIdList**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertyCount** method of an **Image** object to determine the size of the required buffer. The size of the buffer should be the return value of **GetPropertyCount** multiplied by **SIZEOF(PROPID)**.

# <a name="GetPropertyItem"></a>GetPropertyItem (CGpImage)

Gets a specified property item (piece of metadata) from this Image object.

```
FUNCTION GetPropertyItem (BYVAL propId AS PROPID, BYVAL propSize AS UINT, _
   BYVAL buffer AS PropertyItem PTR= AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *propId* | Integer that identifies the property item to be retrieved. |
| *propSize* | Integer that specifies the size, in bytes, of the property item to be retrieved. Call the **GetPropertyItemSize** method to determine the size. |
| *buffer* | Pointer to a **PropertyItem** structure that receives the property item. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPropertyItemSize"></a>GetPropertyItemSize (CGpImage)

Gets the size, in bytes, of a specified property item of this **Image** object.

```
FUNCTION GetPropertyItemSize (BYVAL propId AS PROPID) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *propId* | Integer that identifies the property item. |

# <a name="GetPropertySize"></a>GetPropertySize (CGpImage)

Gets the total size, in bytes, of all the property items stored in this **Image** object. The **GetPropertySize** method also gets the number of property items stored in this Image object.

```
FUNCTION GetPropertySize (BYVAL totalBufferSize AS UINT PTR, BYVAL numProperties AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *totalBufferSize* | Pointer to a UINT that receives the total size, in bytes, of all the property items. |
| *numProperties* | Pointer to a UINT that receives the number of property items. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Windows GDI+ stores an individual piece of metadata in a **PropertyItem** structure. The **GetAllPropertyItems** method returns an array of **PropertyItem** structures. Before you call **GetAllPropertyItems**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertySize** method of an **Image** object to get the size, in bytes, of the required buffer. The **GetPropertySize** method also gives you the number of properties (pieces of metadata) in the image.

# <a name="GetRawFormat"></a>GetRawFormat (CGpImage)

Gets a globally unique identifier ( GUID) that identifies the format of this Image object.

```
FUNCTION GetRawFormat (BYVAL guidformat AS GUID PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *guidformat* | Pointer to a GUID that receives the format identifier. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.
