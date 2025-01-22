# CTextFont2 Class

Class that wraps all the methods of the **ITextFont2** interface.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTORS](#CONSTRUCTORS) | Called when a class variable is created. |
| [DESTRUCTOR](#DESTRUCTOR) | Called automatically when a class variable goes out of scope or is destroyed. |
| [LET](#LET) | Assignment operator. |
| [CAST](#CAST) | Cast operator. |
| [TextDocumentPtr](#TextDocumentPtr) | Returns a pointer to the underlying **ITextFont2** interface. |
| [Attach](#Attach) | Attaches an **ITextFont2** interface pointer to the class. |
| [Detach](#Detach) | Detaches the underlying **ITextFont2** interface pointer from the class. |

### ITextFont2 Interface

In the Text Object Model (TOM), applications access text-range attributes by using a pair of dual interfaces, **ITextFont** and **ITextPara**.

The **ITextFont2** interface extends **ITextFont**, providing the programming equivalent of the Microsoft Word format-font dialog.

| Name       | Description |
| ---------- | ----------- |
| [GetDuplicate](#GetDuplicate) | Gets a duplicate of this text font object. |
| [SetDuplicate](#SetDuplicate) | Sets the character formatting by copying another text font object. |
| [SetDuplicate](#SetDuplicate) |  |
| [CanChange](#CanChange) |  |
| [IsEqual](#IsEqual) |  |
| [Reset](#Reset) |  |
| [GetStyle](#GetStyle) |  |
| [SetStyle](#SetStyle) |  |
| [GetAllCaps](#GetAllCaps) |  |
| [SetAllCaps](#SetAllCaps) |  |
| [GetAnimation](#GetAnimation) |  |
| [SetAnimation](#SetAnimation) |  |
| [GetBackColor](#GetBackColor) |  |
| [SetBackColor](#SetBackColor) |  |
| [GetBold](#GetBold) |  |
| [SetBold](#SetBold) |  |
| [GetEmboss](#GetEmboss) |  |
| [SetEmboss](#SetEmboss) |  |
| [GetForeColor](#GetForeColor) |  |
| [SetForeColor](#SetForeColor) |  |
| [GetHidden](#GetHidden) |  |
| [SetHidden](#SetHidden) |  |
| [GetEngrave](#GetEngrave) |  |
| [SetEngrave](#SetEngrave) |  |
| [GetItalic](#GetItalic) |  |
| [SetItalic](#SetItalic) |  |
| [GetKerning](#GetKerning) |  |
| [SetKerning](#SetKerning) |  |
| [GetLanguageID](#GetLanguageID) |  |
| [SetLanguageID](#SetLanguageID) |  |
| [GetName](#GetName) |  |
| [SetName](#SetName) |  |
| [GetOutline](#GetOutline) |  |
| [SetOutline](#SetOutline) |  |
| [GetPosition](#GetPosition) |  |
| [SetPosition](#SetPosition) |  |
| [GetProtected](#GetProtected) |  |
| [SetProtected](#SetProtected) |  |
| [GetShadow](#GetShadow) |  |
| [SetShadow](#SetShadow) |  |
| [GetSize](#GetSize) |  |
| [SetSize](#SetSize) |  |
| [GetSmallCaps](#GetSmallCaps) |  |
| [SetSmallCaps](#SetSmallCaps) |  |
| [GetSpacing](#GetSpacing) |  |
| [SetSpacing](#SetSpacing) |  |
| [GetStrikeThrough](#GetStrikeThrough) |  |
| [SetStrikeThrough](#SetStrikeThrough) |  |
| [GetSubscript](#GetSubscript) |  |
| [SetSubscript](#SetSubscript) |  |
| [GetSuperscript](#GetSuperscript) |  |
| [SetSuperscript](#SetSuperscript) |  |
| [GetUnderline](#GetUnderline) |  |
| [SetUnderline](#SetUnderline) |  |
| [GetWeight](#GetWeight) |  |
| [SetWeight](#SetWeight) |  |
| [GetCount](#GetCount) |  |
| [GetAutoLigatures](#GetAutoLigatures) |  |
| [SetAutoLigatures](#SetAutoLigatures) |  |
| [GetAutospaceAlpha](#GetAutospaceAlpha) |  |
| [SetAutospaceAlpha](#SetAutospaceAlpha) |  |
| [GetAutospaceNumeric](#GetAutospaceNumeric) |  |
| [SetAutospaceNumeric](#SetAutospaceNumeric) |  |
| [GetAutospaceParens](#GetAutospaceParens) |  |
| [SetAutospaceParens](#SetAutospaceParens) |  |
| [GetCharRep](#GetCharRep) |  |
| [SetCharRep](#SetCharRep) |  |
| [GetCompressionMode](#GetCompressionMode) |  |
| [SetCompressionMode](#SetCompressionMode) |  |
| [GetCookie](#GetCookie) |  |
| [SetCookie](#SetCookie) |  |
| [GetDoubleStrike](#GetDoubleStrike) |  |
| [SetDoubleStrike](#SetDoubleStrike) |  |
| [GetDuplicate2](#GetDuplicate2) |  |
| [SetDuplicate2](#SetDuplicate2) |  |
| [GetLinkType](#GetLinkType) |  |
| [GetMathZone](#GetMathZone) |  |
| [SetMathZone](#SetMathZone) |  |
| [GetModWidthPairs](#GetModWidthPairs) |  |
| [SetModWidthPairs](#SetModWidthPairs) |  |
| [GetModWidthSpace](#GetModWidthSpace) |  |
| [SetModWidthSpace](#SetModWidthSpace) |  |
| [GetOldNumbers](#GetOldNumbers) |  |
| [SetOldNumbers](#SetOldNumbers) |  |
| [GetOverlapping](#GetOverlapping) |  |
| [SetOverlapping](#SetOverlapping) |  |
| [GetPositionSubSuper](#GetPositionSubSuper) |  |
| [SetPositionSubSuper](#SetPositionSubSuper) |  |
| [GetScaling](#GetScaling) |  |
| [SetScaling](#SetScaling) |  |
| [GetSpaceExtension](#GetSpaceExtension) |  |
| [SetSpaceExtension](#SetSpaceExtension) |  |
| [GetUnderlinePositionMode](#GetUnderlinePositionMode) |  |
| [SetUnderlinePositionMode](#SetUnderlinePositionMode) |  |
| [GetEffects](#GetEffects) |  |
| [GetEffects2](#GetEffects2) |  |
| [GetProperty](#GetProperty) |  |
| [GetPropertyInfo](#GetPropertyInfo) |  |
| [IsEqual2](#IsEqual2) |  |
| [SetEffects](#SetEffects) |  |
| [SetEffects2](#SetEffects2) |  |
| [SetProperty](#SetProperty) |  |
