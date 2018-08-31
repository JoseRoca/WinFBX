# CGpImageAttributes Class

An **ImageAttributes** object contains information about how bitmap and metafile colors are manipulated during rendering. An ImageAttributes object maintains several color-adjustment settings, including color-adjustment matrices, grayscale-adjustment matrices, gamma-correction values, color-map tables, and color-threshold values.

**Inherits from**: CGpBase.
**Imclude file**: CGpImageAttributes.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#Constructor) | Creates a new ImageAttributes object.  |
| [ClearBrushRemapTable](#ClearBrushRemapTable) | Clears the brush color-remap table of this ImageAttributes object. |
| [ClearColorKey](#ClearColorKey) | Clears the color key (transparency range) for a specified category. |
| [ClearColorMatrices](#ClearColorMatrices) | Clears the color-adjustment matrix and the grayscale-adjustment matrix for a specified category. |
| [ClearColorMatrix](#ClearColorMatrix) | Clears the color-adjustment matrix for a specified category. |
| [ClearGamma](#ClearGamma) | Disables gamma correction for a specified category. |
| [ClearNoOp](#ClearNoOp) | Clears the NoOp setting for a specified category. |
| [ClearOutputChannel](#ClearOutputChannel) | Clears the cyan-magenta-yellow-black (CMYK) output channel setting for a specified category. |
| [ClearOutputChannelColorProfile](#ClearOutputChannelColorProfile) | Clears the output channel color profile setting for a specified category. |
| [ClearRemapTable](#ClearRemapTable) | Clears the color-remap table for a specified category. |
| [ClearThreshold](#ClearThreshold) | Clears the threshold value for a specified category. |
| [Clone](#Clone) | Copies the contents of the existing ImageAttributes object into a new ImageAttributes object. |
| [GetAdjustedPalette](#GetAdjustedPalette) | Adjusts the colors in a palette according to the adjustment settings of a specified category. |
| [Reset](#Reset) | Clears all color- and grayscale-adjustment settings for a specified category. |
| [SetBrushRemapTable](#SetBrushRemapTable) | Sets the color remap table for the brush category. |
| [SetColorKey](#SetColorKey) | Sets the color key (transparency range) for a specified category. |
| [SetColorMatrices](#SetColorMatrices) | Sets the color-adjustment matrix and the grayscale-adjustment matrix for a specified category. |
| [SetColorMatrix](#SetColorMatrix) | Sets the color-adjustment matrix for a specified category. |
| [SetGamma](#SetGamma) | Sets the gamma value for a specified category. |
| [SetNoOp](#SetNoOp) | Turns off color adjustment for a specified category. |
| [SetOutputChannel](#SetOutputChannel) | Sets the CMYK output channel for a specified category. |
| [SetOutputChannelColorProfile](#SetOutputChannelColorProfile) | Sets the output channel color-profile file for a specified category. |
| [SetRemapTable](#SetRemapTable) | Sets the color-remap table for a specified category. |
| [SetThreshold](#SetThreshold) | Sets the threshold (transparency range) for a specified category. |
| [SetToIdentity](#SetToIdentity) | Sets the color-adjustment matrix of a specified category to identity matrix. |
| [SetWrapMode](#SetWrapMode) | Sets the the wrap mode of this ImageAttributes object. |
