# CGpImage Class

The **CGpImage** class provides methods for loading and saving raster images (bitmaps) and vector images (metafiles). An Image object encapsulates a bitmap or a metafile and stores attributes that you can retrieve by calling various Get methods. You can construct **Image** objects from a variety of file types including BMP, ICON, GIF, JPEG, Exif, PNG, TIFF, WMF, and EMF.

**Inherits from**: CGpBase.
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
