' ########################################################################################
' Microsoft Windows
' File: AfxTOM.bi
' Contents: Interface declarations for the Text Object Model (TOM)
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx/AfxCOM.inc"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' ========================================================================================
' IIDs (Interface identifiers)
' ========================================================================================

CONST AFX_IID_ITextDocument = "{8CC497C0-A1DF-11CE-8098-00AA0047BE5D}"
CONST AFX_IID_ITextDocument2Old = "{01C25500-4268-11D1-883A-3C8B00C10000}" '// RichEdit20
CONST AFX_IID_ITextDocument2 = "{C241F5E0-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST AFX_IID_ITextFont = "{8CC497C3-A1DF-11CE-8098-00AA0047BE5D}"
CONST AFX_IID_ITextPara = "{8CC497C4-A1DF-11CE-8098-00AA0047BE5D}"
CONST AFX_IID_ITextSelection = "{8CC497C1-A1DF-11CE-8098-00AA0047BE5D}"
CONST AFX_IID_ITextRange = "{8CC497C2-A1DF-11CE-8098-00AA0047BE5D}"
CONST AFX_IID_ITextStoryRanges = "{8CC497C5-A1DF-11CE-8098-00AA0047BE5D}"
CONST AFX_IID_ITextMsgFilter = "{A3787420-4267-11D1-883A-3C8B00C10000}"
CONST Afx_IID_ITextFont2 = "{C241F5E3-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST Afx_IID_ITextPara2 = "{C241F5E4-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST Afx_IID_ITextRow = "{C241F5EF-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST Afx_IID_ITextSelection2 = "{C241F5E1-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST Afx_IID_ITextRange2 = "{C241F5E2-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST Afx_IID_ITextDisplays = "{C241F5E2-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST Afx_IID_ITextStrings = "{C241F5E7-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST Afx_IID_ITextStoryRanges2 = "{C241F5E5-7206-11D8-A2C7-00A0D1D6C6B3}"
CONST Afx_IID_ITextStory = "{C241F5F3-7206-11D8-A2C7-00A0D1D6C6B3}"

' ========================================================================================
' Enumerations
' ========================================================================================

ENUM tomConstants
   tomFalse = 0
   tomTrue = -1
   tomUndefined = -9999999
   tomToggle = -9999998
   tomAutoColor = -9999997
   tomDefault = -9999996
   tomSuspend = -9999995
   tomResume = -9999994
   tomApplyNow = 0
   tomApplyLater = 1
   tomTrackParms = 2
   tomCacheParms = 3
   tomApplyTmp	= 4
   tomDisableSmartFont = 8
   tomEnableSmartFont = 9
   tomUsePoints = 10
   tomUseTwips = 11
   tomBackward = &hc0000001
   tomForward = &h3fffffff
   tomMove = 0
   tomExtend = 1
   tomNoSelection = 0
   tomSelectionIP = 1
   tomSelectionNormal = 2
   tomSelectionFrame = 3
   tomSelectionColumn = 4
   tomSelectionRow = 5
   tomSelectionBlock = 6
   tomSelectionInlineShape = 7
   tomSelectionShape = 8
   tomSelStartActive = 1
   tomSelAtEOL = 2
   tomSelOvertype = 4
   tomSelActive = 8
   tomSelReplace = 16
   tomEnd = 0
   tomStart = 32
   tomCollapseEnd = 0
   tomCollapseStart = 1
   tomClientCoord = 256
   tomAllowOffClient = 512
   tomTransform = 1024
   tomObjectArg = 2048
   tomAtEnd = 4096
   tomNone = 0
   tomSingle = 1
   tomWords = 2
   tomDouble = 3
   tomDotted = 4
   tomDash = 5
   tomDashDot = 6
   tomDashDotDot = 7
   tomWave = 8
   tomThick = 9
   tomHair = 10
   tomDoubleWave = 11
   tomHeavyWave = 12
   tomLongDash = 13
   tomThickDash = 14
   tomThickDashDot = 15
   tomThickDashDotDot = 16
   tomThickDotted = 17
   tomThickLongDash = 18
   tomLineSpaceSingle = 0
   tomLineSpace1pt5 = 1
   tomLineSpaceDouble = 2
   tomLineSpaceAtLeast = 3
   tomLineSpaceExactly = 4
   tomLineSpaceMultiple = 5
   tomLineSpacePercent = 6
   tomAlignLeft = 0
   tomAlignCenter = 1
   tomAlignRight = 2
   tomAlignJustify = 3
   tomAlignDecimal = 3
   tomAlignBar = 4
   tomDefaultTab = 5
   tomAlignInterWord = 3
   tomAlignNewspaper = 4
   tomAlignInterLetter = 5
   tomAlignScaled = 6
   tomSpaces = 0
   tomDots = 1
   tomDashes = 2
   tomLines = 3
   tomThickLines = 4
   tomEquals = 5
   tomTabBack = -3
   tomTabNext = -2
   tomTabHere = -1
   tomListNone = 0
   tomListBullet = 1
   tomListNumberAsArabic = 2
   tomListNumberAsLCLetter = 3
   tomListNumberAsUCLetter = 4
   tomListNumberAsLCRoman = 5
   tomListNumberAsUCRoman = 6
   tomListNumberAsSequence = 7
   tomListNumberedCircle = 8
   tomListNumberedBlackCircleWingding = 9
   tomListNumberedWhiteCircleWingding = 10
   tomListNumberedArabicWide = 11
   tomListNumberedChS = 12
   tomListNumberedChT = 13
   tomListNumberedJpnChS = 14
   tomListNumberedJpnKor = 15
   tomListNumberedArabic1 = 16
   tomListNumberedArabic2 = 17
   tomListNumberedHebrew = 18
   tomListNumberedThaiAlpha = 19
   tomListNumberedThaiNum = 20
   tomListNumberedHindiAlpha = 21
   tomListNumberedHindiAlpha1 = 22
   tomListNumberedHindiNum = 23
   tomListParentheses = &h10000
   tomListPeriod = &h20000
   tomListPlain = &h30000
   tomListNoNumber = &h40000
   tomListMinus = &h80000
   tomIgnoreNumberStyle = &h1000000
   tomParaStyleNormal = -1
   tomParaStyleHeading1 = -2
   tomParaStyleHeading2 = -3
   tomParaStyleHeading3 = -4
   tomParaStyleHeading4 = -5
   tomParaStyleHeading5 = -6
   tomParaStyleHeading6 = -7
   tomParaStyleHeading7 = -8
   tomParaStyleHeading8 = -9
   tomParaStyleHeading9 = -10
   tomCharacter = 1
   tomWord = 2
   tomSentence = 3
   tomParagraph = 4
   tomLine = 5
   tomStory = 6
   tomScreen = 7
   tomSection = 8
   tomTableColumn = 9
   tomColumn = 9
   tomRow = 10
   tomWindow = 11
   tomCell = 12
   tomCharFormat = 13
   tomParaFormat = 14
   tomTable = 15
   tomObject = 16
   tomPage = 17
   tomHardParagraph = 18
   tomCluster = 19
   tomInlineObject = 20
   tomInlineObjectArg = 21
   tomLeafLine = 22
   tomLayoutColumn = 23
   tomProcessId = &h40000001
   tomMatchWord = 2
   tomMatchCase = 4
   tomMatchPattern = 8
   tomUnknownStory = 0
   tomMainTextStory = 1
   tomFootnotesStory = 2
   tomEndnotesStory = 3
   tomCommentsStory = 4
   tomTextFrameStory = 5
   tomEvenPagesHeaderStory = 6
   tomPrimaryHeaderStory = 7
   tomEvenPagesFooterStory = 8
   tomPrimaryFooterStory = 9
   tomFirstPageHeaderStory = 10
   tomFirstPageFooterStory = 11
   tomScratchStory = 127
   tomFindStory = 128
   tomReplaceStory = 129
   tomStoryInactive = 0
   tomStoryActiveDisplay = 1
   tomStoryActiveUI = 2
   tomStoryActiveDisplayUI = 3
   tomNoAnimation = 0
   tomLasVegasLights = 1
   tomBlinkingBackground = 2
   tomSparkleText = 3
   tomMarchingBlackAnts = 4
   tomMarchingRedAnts = 5
   tomShimmer = 6
   tomWipeDown = 7
   tomWipeRight = 8
   tomAnimationMax = 8
   tomLowerCase = 0
   tomUpperCase = 1
   tomTitleCase = 2
   tomSentenceCase = 4
   tomToggleCase = 5
   tomReadOnly = &h100
   tomShareDenyRead = &h200
   tomShareDenyWrite = &h400
   tomPasteFile = &h1000
   tomCreateNew = &h10
   tomCreateAlways = &h20
   tomOpenExisting = &h30
   tomOpenAlways = &h40
   tomTruncateExisting = &h50
   tomRTF = &h1
   tomText = &h2
   tomHTML = &h3
   tomWordDocument = &h4
   tomBold = &h80000001
   tomItalic = &h80000002
   tomUnderline = &h80000004
   tomStrikeout = &h80000008
   tomProtected = &h80000010
   tomLink = &h80000020
   tomSmallCaps = &h80000040
   tomAllCaps = &h80000080
   tomHidden = &h80000100
   tomOutline = &h80000200
   tomShadow = &h80000400
   tomEmboss = &h80000800
   tomImprint = &h80001000
   tomDisabled = &h80002000
   tomRevised = &h80004000
   tomSubscriptCF = &h80010000
   tomSuperscriptCF = &h80020000
   tomFontBound = &h80100000
   tomLinkProtected = &h80800000
   tomInlineObjectStart = &h81000000
   tomExtendedChar = &h82000000
   tomAutoBackColor = &h84000000
   tomMathZoneNoBuildUp = &h88000000
   tomMathZone = &h90000000
   tomMathZoneOrdinary = &ha0000000
   tomAutoTextColor = &hc0000000
   tomMathZoneDisplay = &h40000
   tomParaEffectRTL = &h1
   tomParaEffectKeep = &h2
   tomParaEffectKeepNext = &h4
   tomParaEffectPageBreakBefore = &h8
   tomParaEffectNoLineNumber = &h10
   tomParaEffectNoWidowControl = &h20
   tomParaEffectDoNotHyphen = &h40
   tomParaEffectSideBySide = &h80
   tomParaEffectCollapsed = &h100
   tomParaEffectOutlineLevel = &h200
   tomParaEffectBox = &h400
   tomParaEffectTableRowDelimiter = &h1000
   tomParaEffectTable = &h4000
   tomModWidthPairs = &h1
   tomModWidthSpace = &h2
   tomAutoSpaceAlpha = &h4
   tomAutoSpaceNumeric = &h8
   tomAutoSpaceParens = &h10
   tomEmbeddedFont = &h20
   tomDoublestrike = &h40
   tomOverlapping = &h80
   tomNormalCaret = 0
   tomKoreanBlockCaret = &h1
   tomNullCaret = &h2
   tomIncludeInset = &h1
   tomUnicodeBiDi = &h1
   tomMathCFCheck = &h4
   tomUnlink = &h8
   tomUnhide = &h10
   tomCheckTextLimit = &h20
   tomIgnoreCurrentFont = 0
   tomMatchCharRep = &h1
   tomMatchFontSignature = &h2
   tomMatchAscii = &h4
   tomGetHeightOnly = &h8
   tomMatchMathFont = &h10
   tomCharset = &h80000000
   tomCharRepFromLcid = &h40000000
   tomAnsi = 0
   tomEastEurope = 1
   tomCyrillic = 2
   tomGreek = 3
   tomTurkish = 4
   tomHebrew = 5
   tomArabic = 6
   tomBaltic = 7
   tomVietnamese = 8
   tomDefaultCharRep = 9
   tomSymbol = 10
   tomThai = 11
   tomShiftJIS = 12
   tomGB2312 = 13
   tomHangul = 14
   tomBIG5 = 15
   tomPC437 = 16
   tomOEM = 17
   tomMac = 18
   tomArmenian = 19
   tomSyriac = 20
   tomThaana = 21
   tomDevanagari = 22
   tomBengali = 23
   tomGurmukhi = 24
   tomGujarati = 25
   tomOriya = 26
   tomTamil = 27
   tomTelugu = 28
   tomKannada = 29
   tomMalayalam = 30
   tomSinhala = 31
   tomLao = 32
   tomTibetan = 33
   tomMyanmar = 34
   tomGeorgian = 35
   tomJamo = 36
   tomEthiopic = 37
   tomCherokee = 38
   tomAboriginal = 39
   tomOgham = 40
   tomRunic = 41
   tomKhmer = 42
   tomMongolian = 43
   tomBraille = 44
   tomYi = 45
   tomLimbu = 46
   tomTaiLe = 47
   tomNewTaiLue = 48
   tomSylotiNagri = 49
   tomKharoshthi = 50
   tomKayahli = 51
   tomUsymbol = 52
   tomEmoji = 53
   tomGlagolitic = 54
   tomLisu = 55
   tomVai = 56
   tomNKo = 57
   tomOsmanya = 58
   tomPhagsPa = 59
   tomGothic = 60
   tomDeseret = 61
   tomTifinagh = 62
   tomCharRepMax = 63
   tomRE10Mode = &h1
   tomUseAtFont = &h2
   tomTextFlowMask = &hc
   tomTextFlowES = 0
   tomTextFlowSW = &h4
   tomTextFlowWN = &h8
   tomTextFlowNE = &hc
   tomNoIME = &h80000
   tomSelfIME = &h40000
   tomNoUpScroll = &h10000
   tomNoVpScroll = &h40000
   tomNoLink = 0
   tomClientLink = 1
   tomFriendlyLinkName = 2
   tomFriendlyLinkAddress = 3
   tomAutoLinkURL = 4
   tomAutoLinkEmail = 5
   tomAutoLinkPhone = 6
   tomAutoLinkPath = 7
   tomCompressNone = 0
   tomCompressPunctuation = 1
   tomCompressPunctuationAndKana = 2
   tomCompressMax = 2
   tomUnderlinePositionAuto = 0
   tomUnderlinePositionBelow = 1
   tomUnderlinePositionAbove = 2
   tomUnderlinePositionMax = 2
   tomFontAlignmentAuto = 0
   tomFontAlignmentTop = 1
   tomFontAlignmentBaseline = 2
   tomFontAlignmentBottom = 3
   tomFontAlignmentCenter = 4
   tomFontAlignmentMax = 4
   tomRubyBelow = &h80
   tomRubyAlignCenter = 0
   tomRubyAlign010 = 1
   tomRubyAlign121 = 2
   tomRubyAlignLeft = 3
   tomRubyAlignRight = 4
   tomLimitsDefault = 0
   tomLimitsUnderOver = 1
   tomLimitsSubSup = 2
   tomUpperLimitAsSuperScript = 3
   tomLimitsOpposite = 4
   tomShowLLimPlaceHldr = 8
   tomShowULimPlaceHldr = 16
   tomDontGrowWithContent = 64
   tomGrowWithContent = 128
   tomSubSupAlign = 1
   tomLimitAlignMask = 3
   tomLimitAlignCenter = 0
   tomLimitAlignLeft = 1
   tomLimitAlignRight = 2
   tomShowDegPlaceHldr = 8
   tomAlignDefault = 0
   tomAlignMatchAscentDescent = 2
   tomMathVariant = &h20
   tomStyleDefault = 0
   tomStyleScriptScriptCramped = 1
   tomStyleScriptScript = 2
   tomStyleScriptCramped = 3
   tomStyleScript = 4
   tomStyleTextCramped = 5
   tomStyleText = 6
   tomStyleDisplayCramped = 7
   tomStyleDisplay = 8
   tomMathRelSize = &h40
   tomDecDecSize = &hfe
   tomDecSize = &hff
   tomIncSize = ( 1 OR tomMathRelSize ) 
   tomIncIncSize = ( 2 OR tomMathRelSize ) 
   tomGravityUI = 0
   tomGravityBack = 1
   tomGravityFore = 2
   tomGravityIn = 3
   tomGravityOut = 4
   tomGravityBackward = &h20000000
   tomGravityForward = &h40000000
   tomAdjustCRLF = 1
   tomUseCRLF = 2
   tomTextize = 4
   tomAllowFinalEOP = 8
   tomFoldMathAlpha = 16
   tomNoHidden = 32
   tomIncludeNumbering = 64
   tomTranslateTableCell = 128
   tomNoMathZoneBrackets = &h100
   tomConvertMathChar = &h200
   tomNoUCGreekItalic = &h400
   tomAllowMathBold = &h800
   tomLanguageTag = &h1000
   tomConvertRTF = &h2000
   tomApplyRtfDocProps = &h4000
   tomPhantomShow = 1
   tomPhantomZeroWidth = 2
   tomPhantomZeroAscent = 4
   tomPhantomZeroDescent = 8
   tomPhantomTransparent = 16
   tomPhantomASmash = ( tomPhantomShow OR tomPhantomZeroAscent ) 
   tomPhantomDSmash = ( tomPhantomShow OR tomPhantomZeroDescent ) 
   tomPhantomHSmash = ( tomPhantomShow OR tomPhantomZeroWidth ) 
   tomPhantomSmash = ( ( tomPhantomShow OR tomPhantomZeroAscent )  OR tomPhantomZeroDescent ) 
   tomPhantomHorz = ( tomPhantomZeroAscent OR tomPhantomZeroDescent ) 
   tomPhantomVert = tomPhantomZeroWidth
   tomBoxHideTop = 1
   tomBoxHideBottom = 2
   tomBoxHideLeft = 4
   tomBoxHideRight = 8
   tomBoxStrikeH = 16
   tomBoxStrikeV = 32
   tomBoxStrikeTLBR = 64
   tomBoxStrikeBLTR = 128
   tomBoxAlignCenter = 1
   tomSpaceMask = &h1c
   tomSpaceDefault = 0
   tomSpaceUnary = 4
   tomSpaceBinary = 8
   tomSpaceRelational = 12
   tomSpaceSkip = 16
   tomSpaceOrd = 20
   tomSpaceDifferential = 24
   tomSizeText = 32
   tomSizeScript = 64
   tomSizeScriptScript = 96
   tomNoBreak = 128
   tomTransparentForPositioning = 256
   tomTransparentForSpacing = 512
   tomStretchCharBelow = 0
   tomStretchCharAbove = 1
   tomStretchBaseBelow = 2
   tomStretchBaseAbove = 3
   tomMatrixAlignMask = 3
   tomMatrixAlignCenter = 0
   tomMatrixAlignTopRow = 1
   tomMatrixAlignBottomRow = 3
   tomShowMatPlaceHldr = 8
   tomEqArrayLayoutWidth = 1
   tomEqArrayAlignMask = &hc
   tomEqArrayAlignCenter = 0
   tomEqArrayAlignTopRow = 4
   tomEqArrayAlignBottomRow = &hc
   tomMathManualBreakMask = &h7f
   tomMathBreakLeft = &h7d
   tomMathBreakCenter = &h7e
   tomMathBreakRight = &h7f
   tomMathEqAlign = &h80
   tomMathArgShadingStart = &h251
   tomMathArgShadingEnd = &h252
   tomMathObjShadingStart = &h253
   tomMathObjShadingEnd = &h254
   tomFunctionTypeNone = 0
   tomFunctionTypeTakesArg = 1
   tomFunctionTypeTakesLim = 2
   tomFunctionTypeTakesLim2 = 3
   tomFunctionTypeIsLim = 4
   tomMathParaAlignDefault = 0
   tomMathParaAlignCenterGroup = 1
   tomMathParaAlignCenter = 2
   tomMathParaAlignLeft = 3
   tomMathParaAlignRight = 4
   tomMathDispAlignMask = 3
   tomMathDispAlignCenterGroup = 0
   tomMathDispAlignCenter = 1
   tomMathDispAlignLeft = 2
   tomMathDispAlignRight = 3
   tomMathDispIntUnderOver = 4
   tomMathDispFracTeX = 8
   tomMathDispNaryGrow = &h10
   tomMathDocEmptyArgMask = &h60
   tomMathDocEmptyArgAuto = 0
   tomMathDocEmptyArgAlways = &h20
   tomMathDocEmptyArgNever = &h40
   tomMathDocSbSpOpUnchanged = &h80
   tomMathDocDiffMask = &h300
   tomMathDocDiffDefault = 0
   tomMathDocDiffUpright = &h100
   tomMathDocDiffItalic = &h200
   tomMathDocDiffOpenItalic = &h300
   tomMathDispNarySubSup = &h400
   tomMathDispDef = &h800
   tomMathEnableRtl = &h1000
   tomMathBrkBinMask = &h30000
   tomMathBrkBinBefore = 0
   tomMathBrkBinAfter = &h10000
   tomMathBrkBinDup = &h20000
   tomMathBrkBinSubMask = &hc0000
   tomMathBrkBinSubMM = 0
   tomMathBrkBinSubPM = &h40000
   tomMathBrkBinSubMP = &h80000
   tomSelRange = &h255
   tomHstring = &h254
   tomFontPropTeXStyle = &h33c
   tomFontPropAlign = &h33d
   tomFontStretch = &h33e
   tomFontStyle = &h33f
   tomFontStyleUpright = 0
   tomFontStyleOblique = 1
   tomFontStyleItalic = 2
   tomFontStretchDefault = 0
   tomFontStretchUltraCondensed = 1
   tomFontStretchExtraCondensed = 2
   tomFontStretchCondensed = 3
   tomFontStretchSemiCondensed = 4
   tomFontStretchNormal = 5
   tomFontStretchSemiExpanded = 6
   tomFontStretchExpanded = 7
   tomFontStretchExtraExpanded = 8
   tomFontStretchUltraExpanded = 9
   tomFontWeightDefault = 0
   tomFontWeightThin = 100
   tomFontWeightExtraLight = 200
   tomFontWeightLight = 300
   tomFontWeightNormal = 400
   tomFontWeightRegular = 400
   tomFontWeightMedium = 500
   tomFontWeightSemiBold = 600
   tomFontWeightBold = 700
   tomFontWeightExtraBold = 800
   tomFontWeightBlack = 900
   tomFontWeightHeavy = 900
   tomFontWeightExtraBlack = 950
   tomParaPropMathAlign = &h437
   tomDocMathBuild = &h80
   tomMathLMargin = &h81
   tomMathRMargin = &h82
   tomMathWrapIndent = &h83
   tomMathWrapRight = &h84
   tomMathPostSpace = &h86
   tomMathPreSpace = &h85
   tomMathInterSpace = &h87
   tomMathIntraSpace = &h88
   tomCanCopy = &h89
   tomCanRedo = &h8a
   tomCanUndo = &h8b
   tomUndoLimit = &h8c
   tomDocAutoLink = &h8d
   tomEllipsisMode = &h8e
   tomEllipsisState = &h8f
   tomEllipsisNone = 0
   tomEllipsisEnd = 1
   tomEllipsisWord = 3
   tomEllipsisPresent = 1
   tomVTopCell = 1
   tomVLowCell = 2
   tomHStartCell = 4
   tomHContCell = 8
   tomRowUpdate = 1
   tomRowApplyDefault = 0
   tomCellStructureChangeOnly = 1
   tomRowHeightActual = &h80b
END ENUM

ENUM OBJECTTYPE
   tomSimpleText   = 0
   tomRuby   = ( tomSimpleText + 1 ) 
   tomHorzVert  = ( tomRuby + 1 ) 
   tomWarichu   = ( tomHorzVert + 1 ) 
   tomEq  = 9
   tomMath   = 10
   tomAccent = tomMath
   tomBox = ( tomAccent + 1 ) 
   tomBoxedFormula = ( tomBox + 1 ) 
   tomBrackets  = ( tomBoxedFormula + 1 ) 
   tomBracketsWithSeps   = ( tomBrackets + 1 ) 
   tomEquationArray   = ( tomBracketsWithSeps + 1 ) 
   tomFraction  = ( tomEquationArray + 1 ) 
   tomFunctionApply   = ( tomFraction + 1 ) 
   tomLeftSubSup   = ( tomFunctionApply + 1 ) 
   tomLowerLimit   = ( tomLeftSubSup + 1 ) 
   tomMatrix = ( tomLowerLimit + 1 ) 
   tomNary   = ( tomMatrix + 1 ) 
   tomOpChar = ( tomNary + 1 ) 
   tomOverbar   = ( tomOpChar + 1 ) 
   tomPhantom   = ( tomOverbar + 1 ) 
   tomRadical   = ( tomPhantom + 1 ) 
   tomSlashedFraction = ( tomRadical + 1 ) 
   tomStack  = ( tomSlashedFraction + 1 ) 
   tomStretchStack = ( tomStack + 1 ) 
   tomSubscript = ( tomStretchStack + 1 ) 
   tomSubSup = ( tomSubscript + 1 ) 
   tomSuperscript  = ( tomSubSup + 1 ) 
   tomUnderbar  = ( tomSuperscript + 1 ) 
   tomUpperLimit   = ( tomUnderbar + 1 ) 
   tomObjectMax = tomUpperLimit
END ENUM

ENUM MANCODE
   MBOLD  = &h10
   MITAL  = &h20
   MGREEK = &h40
   MROMN  = 0
   MSCRP  = 1
   MFRAK  = 2
   MOPEN  = 3
   MSANS  = 4
   MMONO  = 5
   MMATH  = 6
   MISOL  = 7
   MINIT  = 8
   MTAIL  = 9
   MSTRCH = 10
   MLOOP  = 11
   MOPENA = 12
END ENUM

' ========================================================================================
' Interfaces - Forward references
' ========================================================================================

TYPE ITextMsgFilter AS ITextMsgFilter_

' ========================================================================================
' Dual interfaces - Forward references
' ========================================================================================

TYPE ITextDocument AS ITextDocument_
TYPE ITextDocument2 AS ITextDocument2_
TYPE ITextDocument2Old AS ITextDocument2Old_
TYPE ITextFont AS ITextFont_
TYPE ITextPara AS ITextPara_
TYPE ITextSelection AS ITextSelection_
TYPE ITextRange AS ITextRange_
TYPE ITextStoryRanges AS ITextStoryRanges_
TYPE ITextDisplays AS ITextDisplays_
TYPE ITextFont2 AS ITextFont2_
TYPE ITextPara2 AS ITextPara2_
TYPE ITextRow AS ITextRow_
TYPE ITextSelection2 AS ITextSelection2_
TYPE ITextRange2 AS ITextRange2_
TYPE ITextDisplays AS ITextDisplays_
TYPE ITextStrings AS ITextStrings_
TYPE ITextStoryRanges2 AS ITextStoryRanges2_
TYPE ITextStory AS ITextStory2_

' ========================================================================================
' Interfaces
' ========================================================================================

' ########################################################################################
' Interface name: ITextMsgFilter
' IID: {A3787420-4267-11D1-883A-3C8B00C10000}
' Attributes =  128 [&h00000080] [Nonextensible]
' Inherited interface = IUnknown
' Number of methods = 3
' Note: This interface is undocumented.
' ########################################################################################

#ifndef __ITextMsgFilter_INTERFACE_DEFINED__
#define __ITextMsgFilter_INTERFACE_DEFINED__

TYPE ITextMsgFilterVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextMsgFilter PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextMsgFilter PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextMsgFilter PTR) AS ULONG
   AttachDocument AS FUNCTION (BYVAL this AS ITextMsgFilter PTR, BYVAL hwnd AS HWND, BYVAL pTextDoc AS ITextDocument2 PTR) AS HRESULT
   HandleMessage AS FUNCTION (BYVAL this AS ITextMsgFilter PTR, BYVAL pmsg AS UINT PTR, BYVAL pwparam AS ULONG PTR, BYVAL plparam AS LONG PTR, BYVAL plres AS LONG PTR) AS HRESULT
   AttachMsgFilter AS FUNCTION (BYVAL this AS ITextMsgFilter PTR, BYVAL pMsgFilter AS ITextMsgFilter PTR) AS HRESULT
END TYPE

TYPE ITextMsgFilter_
   lpVtbl AS ITextMsgFilterVTbl PTR
END TYPE

TYPE LPITEXTMSGFILTER AS ITextMsgFilter PTR

#endif

' ########################################################################################
' * Interface name: ITextStory
' IID: {C241F5F3-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __ITextStory_INTERFACE_DEFINED__
#define __ITextStory_INTERFACE_DEFINED__

TYPE ITextStoryVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextStory PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextStory PTR) AS ULONG
   GetActive AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetActive AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL Value AS LONG) AS HRESULT
   GetDisplay AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL ppDisplay AS IUnknown PTR PTR) AS HRESULT
   GetIndex AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetType AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetType AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL Value AS LONG) AS HRESULT
   GetProperty AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL nType AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   GetRange AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   GetText AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL Flags AS LONG, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetFormattedText AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL pUnk AS IUnknown PTR) AS HRESULT
   SetProperty AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   SetText AS FUNCTION (BYVAL this AS ITextStory PTR, BYVAL Flags AS LONG, BYVAL bstr AS BSTR) AS HRESULT
END TYPE

TYPE ITextStory_
   lpVtbl AS ITextStoryVTbl PTR
END TYPE

TYPE LPITEXTSTORY AS ITextStory PTR

#endif

' ========================================================================================
' Dual interfaces
' ========================================================================================

' ########################################################################################
' Interface name: ITextDocument
' IID: {8CC497C0-A1DF-11CE-8098-00AA0047BE5D}
' Attributes =  4288 [&h000010C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 26
' ########################################################################################

#ifndef __ITextDocument_INTERFACE_DEFINED__
#define __ITextDocument_INTERFACE_DEFINED__

TYPE ITextDocumentVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextDocument PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextDocument PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   GetName AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pName AS BSTR PTR) AS HRESULT
   GetSelection AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL ppSel AS ITextSelection PTR PTR) AS HRESULT
   GetStoryCount AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetStoryRanges AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL ppStories AS ITextStoryRanges PTR PTR) AS HRESULT
   GetSaved AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSaved AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL Value AS LONG) AS HRESULT
   GetDefaultTabStop AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetDefaultTabStop AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL Value AS SINGLE) AS HRESULT
   New_ AS FUNCTION (BYVAL this AS ITextDocument PTR) AS HRESULT
   Open AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   Save AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   Freeze AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   Unfreeze AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   BeginEditCollection AS FUNCTION (BYVAL this AS ITextDocument PTR) AS HRESULT
   EndEditCollection AS FUNCTION (BYVAL this AS ITextDocument PTR) AS HRESULT
   Undo AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL Count AS LONG, BYVAL pCount AS LONG PTR) AS HRESULT
   Redo AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL Count AS LONG, BYVAL pCount AS LONG PTR) AS HRESULT
   Range AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   RangeFromPoint AS FUNCTION (BYVAL this AS ITextDocument PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
END TYPE

TYPE ITextDocument_
   lpVtbl AS ITextDocumentVTbl PTR
END TYPE

TYPE LPITEXTDOCUMENT AS ITextDocument PTR

#endif

' ########################################################################################
' Interface name: ITextDocument2
' IID: {01C25500-4268-11D1-883A-3C8B00C10000}
' Attributes =  4288 [&h000010C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 46
' ########################################################################################

#ifndef __ITextDocument2_INTERFACE_DEFINED__
#define __ITextDocument2_INTERFACE_DEFINED__

TYPE ITextDocument2VTbl
   ' // IDispatch methods
   QueryInterface AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextDocument2 PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextDocument2 PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   ' // Inherited ITextDocument interface methods
   GetName AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pName AS BSTR PTR) AS HRESULT
   GetSelection AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppSel AS ITextSelection PTR PTR) AS HRESULT
   GetStoryCount AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetStoryRanges AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppStories AS ITextStoryRanges PTR PTR) AS HRESULT
   GetSaved AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSaved AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetDefaultTabStop AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetDefaultTabStop AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   New_ AS FUNCTION (BYVAL this AS ITextDocument2 PTR) AS HRESULT
   Open AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   Save AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   Freeze AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   Unfreeze AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   BeginEditCollection AS FUNCTION (BYVAL this AS ITextDocument2 PTR) AS HRESULT
   EndEditCollection AS FUNCTION (BYVAL this AS ITextDocument2 PTR) AS HRESULT
   Undo AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Count AS LONG, BYVAL pCount AS LONG PTR) AS HRESULT
   Redo AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Count AS LONG, BYVAL pCount AS LONG PTR) AS HRESULT
   Range AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   RangeFromPoint AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   ' // ITextDocument2 interface methods
   GetCaretType AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCaretType AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetDisplays AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppDisplays AS ITextDisplays PTR PTR) AS HRESULT
   GetDocumentFont AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppFont AS ITextFont2 PTR PTR) AS HRESULT
   SetDocumentFont AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   GetDocumentPara AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppPara AS ITextPara2 PTR PTR) AS HRESULT
   SetDocumentPara AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   GetEastAsianFlags AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pFlags AS LONG PTR) AS HRESULT
   GetGenerator AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetIMEInProgress AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetNotificationMode AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetNotificationMode AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetSelection2 AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppSel AS ITextSelection2 PTR PTR) AS HRESULT
   GetStoryRanges2 AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppStories AS ITextStoryRanges2 PTR PTR) AS HRESULT
   GetTypographyOptions AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pOptions AS LONG PTR) AS HRESULT
   GetVersion AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetWindow AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pHwnd AS __int64 PTR) AS HRESULT
   AttachMsgFilter AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pFilter AS IUnknown PTR) AS HRESULT
   CheckTextLimit AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL cch AS LONG, BYVAL pcch AS LONG PTR) AS HRESULT
   GetCallManager AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppVoid AS IUnknown PTR PTR) AS HRESULT
   GetClientRect AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, BYVAL pTop AS LONG PTR, BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR) AS HRESULT
   GetEffectColor AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Index AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   GetImmContext AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pContext AS __int64 PTR) AS HRESULT
   GetPreferredFont AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL cp AS LONG, BYVAL CharRep AS LONG, _
      BYVAL Options AS LONG, BYVAL curCharRep AS LONG, BYVAL curFontSize AS LONG, BYVAL pbstr AS BSTR PTR, _
      BYVAL pPitchAndFamily AS LONG PTR, BYVAL pNewFontSize AS LONG PTR) AS HRESULT
   GetProperty AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL nType AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   GetStrings AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppStrs AS ITextStrings PTR PTR) AS HRESULT
   Notify AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Notify AS LONG) AS HRESULT
   Range2 AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   RangeFromPoint2 AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   ReleaseCallManager AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pVoid AS IUnknown PTR) AS HRESULT
   ReleaseImmContext AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Context AS __int64) AS HRESULT
   SetEffectColor AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Index AS LONG, BYVAL Value AS ULONG) AS HRESULT
   SetProperty AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   SetTypographyOptions AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Options AS LONG, BYVAL Mask AS LONG) AS HRESULT
   SysBeep AS FUNCTION (BYVAL this AS ITextDocument2 PTR) AS HRESULT
   Update AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Value AS LONG) AS HRESULT
   UpdateWindow AS FUNCTION (BYVAL this AS ITextDocument2 PTR) AS HRESULT
   GetMathProperties AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pOptions AS LONG PTR) AS HRESULT
   SetMathProperties AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pOptions AS LONG, BYVAL Mask AS LONG) AS HRESULT
   GetActiveStory AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppStory AS ITextStory PTR PTR) AS HRESULT
   SetActiveStory AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL pStory AS ITextStory PTR) AS HRESULT
   GetMainStory AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppStory AS ITextStory PTR PTR) AS HRESULT
   GetNewStory AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL ppStory AS ITextStory PTR PTR) AS HRESULT
   GetStory AS FUNCTION (BYVAL this AS ITextDocument2 PTR, BYVAL Index AS LONG, BYVAL ppStory AS ITextStory PTR PTR) AS HRESULT
END TYPE

TYPE ITextDocument2_
   lpVtbl AS ITextDocument2VTbl PTR
END TYPE

TYPE LPITEXTDOCUMENT2 AS ITextDocument2 PTR

#endif

' ########################################################################################
' Interface name: ITextDocument2Old
' IID: {01C25500-4268-11D1-883A-3C8B00C10000}
' Attributes =  4288 [&h000010C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = ITextDocument
' Number of methods = 46
' ########################################################################################

#ifndef __ITextDocument2Old_INTERFACE_DEFINED__
#define __ITextDocument2Old_INTERFACE_DEFINED__

TYPE ITextDocument2OldVTbl
   ' // IDispatch interface methods
   QueryInterface AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextDocument2Old PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextDocument2Old PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   ' // Inherited ITextDocument interface methods
   GetName AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pName AS BSTR PTR) AS HRESULT
   GetSelection AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL ppSel AS ITextSelection PTR PTR) AS HRESULT
   GetStoryCount AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetStoryRanges AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL ppStories AS ITextStoryRanges PTR PTR) AS HRESULT
   GetSaved AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSaved AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Value AS LONG) AS HRESULT
   GetDefaultTabStop AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetDefaultTabStop AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Value AS SINGLE) AS HRESULT
   New_ AS FUNCTION (BYVAL this AS ITextDocument2Old PTR) AS HRESULT
   Open AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   Save AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pVar AS VARIANT PTR, BYVAL Flags AS LONG, BYVAL CodePage AS LONG) AS HRESULT
   Freeze AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   Unfreeze AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   BeginEditCollection AS FUNCTION (BYVAL this AS ITextDocument2Old PTR) AS HRESULT
   EndEditCollection AS FUNCTION (BYVAL this AS ITextDocument2Old PTR) AS HRESULT
   Undo AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Count AS LONG, BYVAL pCount AS LONG PTR) AS HRESULT
   Redo AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Count AS LONG, BYVAL pCount AS LONG PTR) AS HRESULT
   Range AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   RangeFromPoint AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   ' // ITextDocument2Old interface methods
   AttachMsgFilter AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pFilter AS IUnknown PTR) AS HRESULT
   SetEffectColor AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Index AS LONG, BYVAL cr AS ULONG) AS HRESULT
   GetEffectColor AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Index AS LONG, BYVAL pcr AS ULONG PTR) AS HRESULT
   GetCaretType AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pCaretType AS LONG PTR) AS HRESULT
   SetCaretType AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL CaretType AS LONG) AS HRESULT
   GetImmContext AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pContext AS LONG PTR) AS HRESULT
   ReleaseImmContext AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Context AS LONG) AS HRESULT
   GetPreferredFont AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL cp AS LONG, BYVAL CodePage AS LONG, BYVAL Option AS LONG, BYVAL curCodepage AS LONG, BYVAL curFontSize AS LONG, BYVAL pbstr AS BSTR PTR, BYVAL pPitchAndFamily AS LONG PTR, BYVAL pNewFontSize AS LONG PTR) AS HRESULT
   GetNotificationMode AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pMode AS LONG PTR) AS HRESULT
   SetNotificationMode AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Mode AS LONG) AS HRESULT
   GetClientRect AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, BYVAL pTop AS LONG PTR, BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR) AS HRESULT
   GetSelection2 AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL ppSel AS ITextSelection PTR PTR) AS HRESULT
   GetWindow AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL phWnd AS LONG PTR) AS HRESULT
   GetFEFlags AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL pFlags AS LONG PTR) AS HRESULT
   UpdateWindow AS FUNCTION (BYVAL this AS ITextDocument2Old PTR) AS HRESULT
   CheckTextLimit AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL cch AS LONG, BYVAL pcch AS LONG PTR) AS HRESULT
   IMEInProgress AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Mode AS LONG) AS HRESULT
   SysBeep AS FUNCTION (BYVAL this AS ITextDocument2Old PTR) AS HRESULT
   Update AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Mode AS LONG) AS HRESULT
   Notify AS FUNCTION (BYVAL this AS ITextDocument2Old PTR, BYVAL Notify AS LONG) AS HRESULT
END TYPE

TYPE ITextDocument2Old_
   lpVtbl AS ITextDocument2OldVTbl PTR
END TYPE

TYPE LPITextDocument2Old AS ITextDocument2Old PTR

#endif

' ########################################################################################
' * Interface name: ITextFont
' IID: {8CC497C3-A1DF-11CE-8098-00AA0047BE5D}
' Attributes =  4288 [&h000010C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 62
' ########################################################################################

#ifndef __ITextFont_INTERFACE_DEFINED__
#define __ITextFont_INTERFACE_DEFINED__

TYPE ITextFontVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextFont PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextFont PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   GetDuplicate AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL ppFont AS ITextFont PTR PTR) AS HRESULT
   SetDuplicate AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pFont AS ITextFont PTR) AS HRESULT
   CanChange AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pFont AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Reset AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetStyle AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetStyle AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetAllCaps AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAllCaps AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetAnimation AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAnimation AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetBackColor AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetBackColor AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetBold AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetBold AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   SetEmboss AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetForeColor AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetForeColor AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetHidden AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetHidden AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetEngrave AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetEngrave AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetItalic AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetItalic AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetKerning AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetKerning AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetLanguageID AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetLanguageID AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetName AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetName AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL bstr AS BSTR) AS HRESULT
   GetOutline AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetOutline AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetPosition AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetPosition AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetProtected AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetProtected AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetShadow AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetShadow AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetSize AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetSize AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetSmallCaps AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSmallCaps AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetSpacing AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetSpacing AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetStrikeThrough AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetStrikeThrough AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetSubscript AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSubscript AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetSuperscript AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSuperscript AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetUnderline AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetUnderline AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT
   GetWeight AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetWeight AS FUNCTION (BYVAL this AS ITextFont PTR, BYVAL Value AS LONG) AS HRESULT

END TYPE

TYPE ITextFont_
   lpVtbl AS ITextFontVTbl PTR
END TYPE

TYPE LPITEXTFONT AS ITextFont PTR

#endif

' ########################################################################################
' * Interface name: ITextFont2
' IID: {C241F5E3-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = ITextFont
' ########################################################################################

#ifndef __ITextFont2_INTERFACE_DEFINED__
#define __ITextFont2_INTERFACE_DEFINED__

TYPE ITextFont2VTbl
   ' // IDispatch interface methods
   QueryInterface AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextFont2 PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextFont2 PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   ' // Inherited ITextFont interface methods
   GetDuplicate AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL ppFont AS ITextFont2 PTR PTR) AS HRESULT
   SetDuplicate AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   CanChange AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pFont AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Reset AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetStyle AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetStyle AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetAllCaps AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAllCaps AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetAnimation AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAnimation AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetBackColor AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetBackColor AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetBold AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetBold AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetEmboss AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetEmboss AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetForeColor AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetForeColor AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetHidden AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetHidden AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetEngrave AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetEngrave AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetItalic AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetItalic AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetKerning AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetKerning AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetLanguageID AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetLanguageID AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetName AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetName AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL bstr AS BSTR) AS HRESULT
   GetOutline AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetOutline AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetPosition AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetPosition AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetProtected AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetProtected AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetShadow AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetShadow AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetSize AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetSize AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetSmallCaps AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSmallCaps AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetSpacing AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetSpacing AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetStrikeThrough AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetStrikeThrough AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetSubscript AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSubscript AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetSuperscript AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSuperscript AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetUnderline AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetUnderline AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetWeight AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetWeight AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   ' // ITextFont2 interface methods
   GetCount AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetAutoLigatures AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAutoLigatures AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetAutospaceAlpha AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAutospaceAlpha AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetAutospaceNumeric AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAutospaceNumeric AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetAutospaceParens AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAutospaceParens AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetCharRep AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCharRep AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetCompressionMode AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCompressionMode AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetCookie AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCookie AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetDoubleStrike AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetDoubleStrike AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetDuplicate2 AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL ppFont AS ITextFont2 PTR PTR) AS HRESULT
   SetDuplicate2 AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   GetLinkType AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetMathZone AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetMathZone AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetModWidthPairs AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetModWidthPairs AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetModWidthSpace AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetModWidthSpace AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetOldNumbers AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetOldNumbers AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetOverlapping AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetOverlapping AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetPositionSubSuper AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetPositionSubSuper AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetScaling AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetScaling AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetSpaceExtension AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSpaceExtension AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetUnderlinePositionMode AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetUnderlinePositionMode AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetEffects AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   GetEffects2 AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   GetProperty AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL nType AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   GetPropertyInfo AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Index AS LONG, BYVAL pType AS LONG PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual2 AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL pFont AS ITextFont2 PTR, BYVAL pB AS LONG PTR) AS HRESULT
   SetEffects AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   SetEffects2 AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   SetProperty AS FUNCTION (BYVAL this AS ITextFont2 PTR, BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
END TYPE

TYPE ITextFont2_
   lpVtbl AS ITextFont2VTbl PTR
END TYPE

TYPE LPITEXTFONT2 AS ITextFont2 PTR

#endif

' ########################################################################################
' * Interface name: ITextPara
' IID: {8CC497C4-A1DF-11CE-8098-00AA0047BE5D}
' Attributes =  4288 [&h000010C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 55
' ########################################################################################

#ifndef __ITextPara_INTERFACE_DEFINED__
#define __ITextPara_INTERFACE_DEFINED__

TYPE ITextParaVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextPara PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextPara PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   GetDuplicate AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL ppPara AS ITextPara PTR PTR) AS HRESULT
   SetDuplicate AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pPara AS ITextPara PTR) AS HRESULT
   CanChange AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pPara AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Reset AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetStyle AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetStyle AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetAlignment AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAlignment AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetHyphenation AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetHyphenation AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetFirstLineIndent AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   GetKeepTogether AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetKeepTogether AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetKeepWithNext AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetKeepWithNext AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetLeftIndent AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   GetLineSpacing AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   GetLineSpacingRule AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetListAlignment AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetListAlignment AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetListLevelIndex AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetListLevelIndex AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetListStart AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetListStart AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetListTab AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetListTab AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetListType AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetListType AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetNoLineNumber AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetNoLineNumber AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetPageBreakBefore AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetPageBreakBefore AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetRightIndent AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetRightIndent AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS SINGLE) AS HRESULT
   SetIndents AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL First AS SINGLE, BYVAL Left AS SINGLE, BYVAL Right AS SINGLE) AS HRESULT
   SetLineSpacing AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Rule AS LONG, BYVAL Spacing AS SINGLE) AS HRESULT
   GetSpaceAfter AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetSpaceAfter AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetSpaceBefore AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetSpaceBefore AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetWidowControl AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetWidowControl AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL Value AS LONG) AS HRESULT
   GetTabCount AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   AddTab AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL tbPos AS SINGLE, BYVAL tbAlign AS LONG, BYVAL tbLeader AS LONG) AS HRESULT
   ClearAllTabs AS FUNCTION (BYVAL this AS ITextPara PTR) AS HRESULT
   DeleteTab AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL tbPos AS SINGLE) AS HRESULT
   GetTab AS FUNCTION (BYVAL this AS ITextPara PTR, BYVAL iTab AS LONG, BYVAL ptbPos AS SINGLE PTR, BYVAL ptbAlign AS LONG PTR, BYVAL ptbLeader AS LONG PTR) AS HRESULT
END TYPE


TYPE ITextPara_
   lpVtbl AS ITextParaVTbl PTR
END TYPE

TYPE LPITEXTPARA AS ITextPara PTR

#endif

' ########################################################################################
' * Interface name: ITextPara2
' IID: {C241F5E4-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = ITextPara
' ########################################################################################

#ifndef __ITextPara2_INTERFACE_DEFINED__
#define __ITextPara2_INTERFACE_DEFINED__

TYPE ITextPara2VTbl
   ' // IDispatch interface methods
   QueryInterface AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextPara2 PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextPara2 PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   ' // Inherited ITextPara interface methods
   GetDuplicate AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL ppPara AS ITextPara2 PTR PTR) AS HRESULT
   SetDuplicate AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   CanChange AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pPara AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Reset AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetStyle AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetStyle AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetAlignment AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAlignment AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetHyphenation AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetHyphenation AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetFirstLineIndent AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   GetKeepTogether AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetKeepTogether AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetKeepWithNext AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetKeepWithNext AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetLeftIndent AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   GetLineSpacing AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   GetLineSpacingRule AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetListAlignment AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetListAlignment AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetListLevelIndex AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetListLevelIndex AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetListStart AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetListStart AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetListTab AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetListTab AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetListType AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetListType AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetNoLineNumber AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetNoLineNumber AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetPageBreakBefore AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetPageBreakBefore AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetRightIndent AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetRightIndent AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   SetIndents AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL First AS SINGLE, BYVAL Left AS SINGLE, BYVAL Right AS SINGLE) AS HRESULT
   SetLineSpacing AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Rule AS LONG, BYVAL Spacing AS SINGLE) AS HRESULT
   GetSpaceAfter AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetSpaceAfter AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetSpaceBefore AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS SINGLE PTR) AS HRESULT
   SetSpaceBefore AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS SINGLE) AS HRESULT
   GetWidowControl AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetWidowControl AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetTabCount AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   AddTab AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL tbPos AS SINGLE, BYVAL tbAlign AS LONG, BYVAL tbLeader AS LONG) AS HRESULT
   ClearAllTabs AS FUNCTION (BYVAL this AS ITextPara2 PTR) AS HRESULT
   DeleteTab AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL tbPos AS SINGLE) AS HRESULT
   GetTab AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL iTab AS LONG, BYVAL ptbPos AS SINGLE PTR, BYVAL ptbAlign AS LONG PTR, BYVAL ptbLeader AS LONG PTR) AS HRESULT
   ' // ITextPara2 interface methods
   GetBorders AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL ppBorders AS IUnknown PTR PTR) AS HRESULT
   GetDuplicate2 AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL ppPara AS ITextPara2 PTR PTR) AS HRESULT
   SetDuplicate2 AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   GetFontAlignment AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetFontAlignment AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetHangingPunctuation AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetHangingPunctuation AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetSnapToGrid AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetSnapToGrid AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetTrimPunctuationAtStart AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetTrimPunctuationAtStart AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetEffects AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pValue AS LONG PTR, BYVAL pMask AS LONG PTR) AS HRESULT
   GetProperty AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL nType AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual2 AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL pPara AS ITextPara2 PTR, BYVAL pB AS LONG PTR) AS HRESULT
   SetEffects AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL Value AS LONG, BYVAL Mask AS LONG) AS HRESULT
   SetProperty AS FUNCTION (BYVAL this AS ITextPara2 PTR, BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
END TYPE


TYPE ITextPara2_
   lpVtbl AS ITextPara2VTbl PTR
END TYPE

TYPE LPITEXTPARA2 AS ITextPara2 PTR

#endif

' ########################################################################################
' * Interface name: ITextRange
' IID: {8CC497C2-A1DF-11CE-8098-00AA0047BE5D}
' Attributes =  4288 [&h000010C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 58
' ########################################################################################

#ifndef __ITextRange_INTERFACE_DEFINED__
#define __ITextRange_INTERFACE_DEFINED__

TYPE ITextRangeVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextRange PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextRange PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   GetText AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetText AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL bstr AS BSTR) AS HRESULT
   GetChar AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pChar AS LONG PTR) AS HRESULT
   SetChar AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Char AS LONG) AS HRESULT
   GetDuplicate AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   GetFormattedText AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   SetFormattedText AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pRange AS ITextRange PTR) AS HRESULT
   GetStart AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pcpFirst AS LONG PTR) AS HRESULT
   SetStart AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL cpFirst AS LONG) AS HRESULT
   GetEnd AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   SetEnd AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL cpLim AS LONG) AS HRESULT
   GetFont AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL ppFont AS ITextFont PTR PTR) AS HRESULT
   SetFont AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pFont AS ITextFont PTR) AS HRESULT
   GetPara AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL ppPara AS ITextPara PTR PTR) AS HRESULT
   SetPara AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pPara AS ITextPara PTR) AS HRESULT
   GetStoryLength AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetStoryType AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Collapse AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL bStart AS LONG) AS HRESULT
   Expand AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   GetIndex AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL pIndex AS LONG PTR) AS HRESULT
   SetIndex AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL Index AS LONG, BYVAL Extend AS LONG) AS HRESULT
   SetRange AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   InRange AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pRange AS ITextRange PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   InStory AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pRange AS ITextRange PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pRange AS ITextRange PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Select AS FUNCTION (BYVAL this AS ITextRange PTR) AS HRESULT
   StartOf AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   EndOf AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   Move AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStart AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEnd AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveWhile AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStartWhile AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEndWhile AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveUntil AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStartUntil AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEndUntil AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   FindText AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL bstr AS BSTR, BYVAL cch AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   FindTextStart AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   FindTextEnd AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   Delete_ AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   Cut AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   Copy AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   Paste AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS HRESULT
   CanPaste AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   CanEdit AS FUNCTION (BYVAL this AS ITextRange PTR, BYVal pValue AS LONG PTR) AS HRESULT
   ChangeCase AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL nType AS LONG) AS HRESULT
   GetPoint AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL nType AS LONG, BYVAL px AS LONG PTR, BYVAL py AS LONG PTR) AS HRESULT
   SetPoint AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL Extend AS LONG) AS HRESULT
   ScrollIntoView AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL Value AS LONG) AS HRESULT
   GetEmbeddedObject AS FUNCTION (BYVAL this AS ITextRange PTR, BYVAL ppObject AS IUnknown PTR PTR) AS HRESULT
END TYPE


TYPE ITextRange_
   lpVtbl AS ITextRangeVTbl PTR
END TYPE

TYPE LPITEXTRANGE AS ITextRange PTR

#endif

' ########################################################################################
' * Interface name: ITextRange2
' IID: {C241F5E2-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = ITextRange
' ########################################################################################

#ifndef __ITextRange2_INTERFACE_DEFINED__
#define __ITextRange2_INTERFACE_DEFINED__

TYPE ITextRange2VTbl
   ' // IDispatch interface methods
   QueryInterface AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextRange2 PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextRange2 PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   ' // ITextRange interface methods
   GetText AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetText AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL bstr AS BSTR) AS HRESULT
   GetChar AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pChar AS LONG PTR) AS HRESULT
   SetChar AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Char AS LONG) AS HRESULT
   GetDuplicate AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   GetFormattedText AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   SetFormattedText AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   GetStart AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pcpFirst AS LONG PTR) AS HRESULT
   SetStart AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL cpFirst AS LONG) AS HRESULT
   GetEnd AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   SetEnd AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL cpLim AS LONG) AS HRESULT
   GetFont AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppFont AS ITextFont PTR PTR) AS HRESULT
   SetFont AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pFont AS ITextFont PTR) AS HRESULT
   GetPara AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppPara AS ITextPara PTR PTR) AS HRESULT
   SetPara AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pPara AS ITextPara PTR) AS HRESULT
   GetStoryLength AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetStoryType AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Collapse AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL bStart AS LONG) AS HRESULT
   Expand AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   GetIndex AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL pIndex AS LONG PTR) AS HRESULT
   SetIndex AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Index AS LONG, BYVAL Extend AS LONG) AS HRESULT
   SetRange AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   InRange AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pRange AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   InStory AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pRange AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pRange AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Select AS FUNCTION (BYVAL this AS ITextRange2 PTR) AS HRESULT
   StartOf AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   EndOf AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   Move AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStart AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEnd AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveWhile AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStartWhile AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEndWhile AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveUntil AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStartUntil AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEndUntil AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   FindText AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL bstr AS BSTR, BYVAL cch AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   FindTextStart AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   FindTextEnd AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   Delete_ AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   Cut AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   Copy AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   Paste AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS HRESULT
   CanPaste AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   CanEdit AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVal pValue AS LONG PTR) AS HRESULT
   ChangeCase AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL nType AS LONG) AS HRESULT
   GetPoint AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL nType AS LONG, BYVAL px AS LONG PTR, BYVAL py AS LONG PTR) AS HRESULT
   SetPoint AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL Extend AS LONG) AS HRESULT
   ScrollIntoView AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetEmbeddedObject AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppObject AS IUnknown PTR PTR) AS HRESULT
   ' // ITextRange2 interface methods
   GetFlags AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pFlags AS LONG PTR) AS HRESULT
   SetFlags AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Flags AS LONG) AS HRESULT
   GetType AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pType AS LONG PTR) AS HRESULT
   MoveLeft AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveRight AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveUp AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveDown AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   HomeKey AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   EndKey AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   TypeText AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL bstr as BSTR) AS HRESULT
   GetCch AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pcch as LONG PTR) AS HRESULT
   GetCells AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppCells AS IUnknown PTR PTR) AS HRESULT
   GetColumn AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppColumn AS IUnknown PTR PTR) AS HRESULT
   GetCount AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetDuplicate2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   GetFont2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppFont AS ITextFont2 PTR PTR) AS HRESULT
   SetFont2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   GetFormattedText2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   SetFormattedText2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   GetGravity AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetGravity AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetPara2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppPara AS ITextPara2 PTR PTR) AS HRESULT
   SetPara2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   GetRow AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppRow AS ITextRow PTR PTR) AS HRESULT
   GetStartPara AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetTable AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL ppTable AS IUnknown PTR PTR) AS HRESULT
   GetURL AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetURL AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL bstr AS BSTR) AS HRESULT
   AddSubrange AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL cp1 AS LONG, BYVAL cp2 AS LONG, BYVAL Activate AS LONG) AS HRESULT
   BuildUpMath AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Flags AS LONG) AS HRESULT
   DeleteSubrange AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL cpFirst AS LONG, BYVAL cpLim AS LONG) AS HRESULT
   Find AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pRange AS ITextRange2 PTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   GetChar2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pChar AS LONG PTR, BYVAL Offset AS LONG) AS HRESULT
   GetDropCap AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pcLine AS LONG PTR, BYVAL pPosition AS LONG PTR) AS HRESULT
   GetInlineObject AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL pType AS LONG PTR, BYVAL pAlign AS LONG PTR, BYVAL pChar AS LONG PTR, BYVAL pChar1 AS LONG PTR, _
      BYVAL pChar2 AS LONG PTR, BYVAL pCount AS LONG PTR,  BYVAL pTeXStyle AS LONG PTR, BYVAL pcCol AS LONG PTR, BYVAL pLevel AS LONG PTR) AS HRESULT
   GetProperty AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL nType AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   GetRect AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, _
      BYVAL pTop AS LONG PTR, BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR, BYVAL pHit AS LONG PTR) AS HRESULT
   GetSubrange AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL iSubrange AS LONG, BYVAL pcpFirst AS LONG PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   GetText2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Flags AS LONG, BYVAL pbstr AS BSTR PTR) AS HRESULT
   HexToUnicode AS FUNCTION (BYVAL this AS ITextRange2 PTR) AS HRESULT
   InsertTable AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL cCol AS LONG, BYVAL cRow AS LONG, BYVAL AutoFit AS LONG) AS HRESULT
   Linearize AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Flags AS LONG) AS HRESULT
   SetActiveSubrange AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   SetDropCap AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL cLine AS LONG, BYVAL Position AS LONG) AS HRESULT
   SetProperty AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   SetText2 AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL Flags AS LONG, BYVAL bstr AS BSTR) AS HRESULT
   UnicodeToHex AS FUNCTION (BYVAL this AS ITextRange2 PTR) AS HRESULT
   SetInlineObject AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL nType AS Long, BYVAL Align AS LONG, BYVAL Char AS LONG, _
      BYVAL Char1 AS LONG, BYVAL Char2 AS LONG, BYVAL Count AS LONG, BYVAL TeXStyle AS LONG, BYVAL cCol AS LONG) AS HRESULT
   GetMathFunctionType AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL bstr AS BSTR, BYVAL pValue AS LONG PTR) AS HRESULT
   InsertImage AS FUNCTION (BYVAL this AS ITextRange2 PTR, BYVAL width_ AS LONG, BYVAL height AS LONG, BYVAL ascent AS LONG, _
      BYVAL nType AS LONG, BYVAL bstrAltText AS BSTR, BYVAL pStream AS IStream PTR) AS HRESULT
END TYPE

TYPE ITextRange2_
   lpVtbl AS ITextRange2VTbl PTR
END TYPE

TYPE LPITEXTRANGE2 AS ITextRange2 PTR

#endif

' ########################################################################################
' * Interface name: ITextSelection
' IID: {8CC497C1-A1DF-11CE-8098-00AA0047BE5D}
' Attributes =  4288 [&h000010C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = ITextRange
' Number of methods = 68
' ########################################################################################

#ifndef __ITextSelection_INTERFACE_DEFINED__
#define __ITextSelection_INTERFACE_DEFINED__

TYPE ITextSelectionVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextSelection PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextSelection PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   GetText AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetText AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL bstr AS BSTR) AS HRESULT
   GetChar AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pChar AS LONG PTR) AS HRESULT
   SetChar AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Char AS LONG) AS HRESULT
   GetDuplicate AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   GetFormattedText AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   SetFormattedText AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pRange AS ITextRange PTR) AS HRESULT
   GetStart AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pcpFirst AS LONG PTR) AS HRESULT
   SetStart AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL cpFirst AS LONG) AS HRESULT
   GetEnd AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   SetEnd AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL cpLim AS LONG) AS HRESULT
   GetFont AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL ppFont AS ITextFont PTR PTR) AS HRESULT
   SetFont AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pFont AS ITextFont PTR) AS HRESULT
   GetPara AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL ppPara AS ITextPara PTR PTR) AS HRESULT
   SetPara AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pPara AS ITextPara PTR) AS HRESULT
   GetStoryLength AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetStoryType AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Collapse AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL bStart AS LONG) AS HRESULT
   Expand AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   GetIndex AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL pIndex AS LONG PTR) AS HRESULT
   SetIndex AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Index AS LONG, BYVAL Extend AS LONG) AS HRESULT
   SetRange AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   InRange AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pRange AS ITextRange PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   InStory AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pRange AS ITextRange PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pRange AS ITextRange PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Select AS FUNCTION (BYVAL this AS ITextSelection PTR) AS HRESULT
   StartOf AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   EndOf AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   Move AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStart AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEnd AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveWhile AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStartWhile AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEndWhile AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveUntil AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStartUntil AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEndUntil AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   FindText AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   FindTextStart AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   FindTextEnd AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   Delete_ AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   Cut AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   Copy AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   Paste AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS HRESULT
   CanPaste AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   CanEdit AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   ChangeCase AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL nType AS LONG) AS HRESULT
   GetPoint AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL nType AS LONG, BYVAL px AS LONG PTR, BYVAL py AS LONG PTR) AS HRESULT
   SetPoint AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL Extend AS LONG) AS HRESULT
   ScrollIntoView AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Value AS LONG) AS HRESULT
   GetEmbeddedObject AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL ppObject AS IUnknown PTR PTR) AS HRESULT
   GetFlags AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pFlags AS LONG PTR) AS HRESULT
   SetFlags AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Flags AS LONG) AS HRESULT
   GetType AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL pType AS LONG PTR) AS HRESULT
   MoveLeft AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveRight AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveUp AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveDown AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   HomeKey AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   EndKey AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   TypeText AS FUNCTION (BYVAL this AS ITextSelection PTR, BYVAL bstr AS BSTR) AS HRESULT
END TYPE

TYPE ITextSelection_
   lpVtbl AS ITextSelectionVTbl PTR
END TYPE

TYPE LPITEXTSELECTION AS ITextSelection PTR

#endif

' ########################################################################################
' * Interface name: ITextSelection2
' IID: {C241F5E1-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = ITextSelection
' Currently, this interface contains no methods other than those inherited from ITextRange2.
' ########################################################################################

#ifndef __ITextSelection2_INTERFACE_DEFINED__
#define __ITextSelection2_INTERFACE_DEFINED__

TYPE ITextSelection2VTbl
   ' // IDispatch interface methods
   QueryInterface AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextSelection2 PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextSelection2 PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   ' // Inherited ITextSelection interface methods
   GetText AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetText AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL bstr AS BSTR) AS HRESULT
   GetChar AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pChar AS LONG PTR) AS HRESULT
   SetChar AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Char AS LONG) AS HRESULT
   GetDuplicate AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   GetFormattedText AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   SetFormattedText AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   GetStart AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pcpFirst AS LONG PTR) AS HRESULT
   SetStart AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL cpFirst AS LONG) AS HRESULT
   GetEnd AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   SetEnd AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL cpLim AS LONG) AS HRESULT
   GetFont AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppFont AS ITextFont PTR PTR) AS HRESULT
   SetFont AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pFont AS ITextFont PTR) AS HRESULT
   GetPara AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppPara AS ITextPara PTR PTR) AS HRESULT
   SetPara AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pPara AS ITextPara PTR) AS HRESULT
   GetStoryLength AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetStoryType AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Collapse AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL bStart AS LONG) AS HRESULT
   Expand AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   GetIndex AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL pIndex AS LONG PTR) AS HRESULT
   SetIndex AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Index AS LONG, BYVAL Extend AS LONG) AS HRESULT
   SetRange AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   InRange AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pRange AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   InStory AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pRange AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pRange AS ITextRange2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   Select AS FUNCTION (BYVAL this AS ITextSelection2 PTR) AS HRESULT
   StartOf AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   EndOf AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   Move AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStart AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEnd AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveWhile AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStartWhile AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEndWhile AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveUntil AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveStartUntil AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveEndUntil AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Cset AS VARIANT PTR, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   FindText AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   FindTextStart AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   FindTextEnd AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL bstr AS BSTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pLength AS LONG PTR) AS HRESULT
   Delete_ AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   Cut AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   Copy AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   Paste AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG) AS HRESULT
   CanPaste AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pVar AS VARIANT PTR, BYVAL Format AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   CanEdit AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   ChangeCase AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL nType AS LONG) AS HRESULT
   GetPoint AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL nType AS LONG, BYVAL px AS LONG PTR, BYVAL py AS LONG PTR) AS HRESULT
   SetPoint AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nType AS LONG, BYVAL Extend AS LONG) AS HRESULT
   ScrollIntoView AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetEmbeddedObject AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppObject AS IUnknown PTR PTR) AS HRESULT
   GetFlags AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pFlags AS LONG PTR) AS HRESULT
   SetFlags AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Flags AS LONG) AS HRESULT
   GetType AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pType AS LONG PTR) AS HRESULT
   MoveLeft AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveRight AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveUp AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   MoveDown AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Count AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   HomeKey AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   EndKey AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Unit AS LONG, BYVAL Extend AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   TypeText AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL bstr AS BSTR) AS HRESULT
   ' // ITextSelection2 interface methods
   GetCch AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pcch AS LONG PTR) AS HRESULT
   GetCells AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppCells AS IUnknown PTR PTR) AS HRESULT
   GetColumn AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppColumn AS IUnknown PTR PTR) AS HRESULT
   GetCount AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   GetDuplicate2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   GetFont2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppFont AS ITextFont2 PTR PTR) AS HRESULT
   SetFont2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pFont AS ITextFont2 PTR) AS HRESULT
   GetFormattedText2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
   SetFormattedText2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   GetGravity AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetGravity AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Value AS LONG) AS HRESULT
   GetPara2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppPara AS ITextPara2 PTR PTR) AS HRESULT
   SetPara2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pPara AS ITextPara2 PTR) AS HRESULT
   GetRow AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppRow AS ITextRow PTR PTR) AS HRESULT
   GetStartPara AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetTable AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL ppTable AS IUnknown PTR PTR) AS HRESULT
   GetURL AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pbstr AS BSTR PTR) AS HRESULT
   SetURL AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL bstr AS BSTR) AS HRESULT
   AddSubrange AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL cp1 AS LONG, BYVAL cp2 AS LONG, BYVAL Activate AS LONG) AS HRESULT
   BuildUpMath AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Flags AS LONG) AS HRESULT
   DeleteSubrange AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL cpFirst AS LONG, BYVAL cpLim AS LONG) AS HRESULT
   Find AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pRange AS ITextRange2 PTR, BYVAL Count AS LONG, BYVAL Flags AS LONG, BYVAL pDelta AS LONG PTR) AS HRESULT
   GetChar2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pChar AS LONG PTR, BYVAL Offset AS LONG) AS HRESULT
   GetDropCap AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pcLine AS LONG PTR, BYVAL pPosition AS LONG PTR) AS HRESULT
   GetInlineObject AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL pType AS LONG PTR, BYVAL pAlign AS LONG PTR, _
      BYVAL pChar AS LONG PTR, BYVAL pChar1 AS LONG PTR, BYVAL pChar2 AS LONG PTR, BYVAL pCount AS LONG PTR, _
      BYVAL pTeXStyle AS LONG PTR, BYVAL pcCol AS LONG PTR, BYVAL pLevel AS LONG PTR) AS HRESULT
   GetProperty AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL nType AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   GetRect AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL nType AS LONG, BYVAL pLeft AS LONG PTR, BYVAL pTop AS LONG PTR, _
      BYVAL pRight AS LONG PTR, BYVAL pBottom AS LONG PTR, BYVAL pHit AS LONG PTR) AS HRESULT
   GetSubrange AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL iSubrange AS LONG, BYVAL pcpFirst As LONG PTR, BYVAL pcpLim AS LONG PTR) AS HRESULT
   GetText2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Flags AS LONG, BYVAL pbstr AS BSTR PTR) AS HRESULT
   HexToUnicode AS FUNCTION (BYVAL this AS ITextSelection2 PTR) AS HRESULT
   InsertTable AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL cCol AS LONG, BYVAL cRow AS LONG, BYVAL AutoFit AS LONG) AS HRESULT
   Linearize AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Flags AS LONG) AS HRESULT
   SetActiveSubrange AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL cpAnchor AS LONG, BYVAL cpActive AS LONG) AS HRESULT
   SetDropCap AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL cLine AS LONG, BYVAL Position AS LONG) AS HRESULT
   SetProperty AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
   SetText2 AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL Flags AS LONG, BYVAL bstr AS BSTR) AS HRESULT
   UnicodeToHex AS FUNCTION (BYVAL this AS ITextSelection2 PTR) AS HRESULT
   SetInlineObject AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL nType AS LONG, BYVAL Align AS LONG, _
      BYVAL Char AS LONG, BYVAL Char1 AS LONG, BYVAL Char2 AS LONG, BYVAL Count AS LONG, BYVAL TeXStyle AS LONG, BYVAL cCol AS LONG) AS HRESULT
   GetMathFunctionType AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL bstr AS BSTR, BYVAL pValue AS LONG PTR) AS HRESULT
   InsertImage AS FUNCTION (BYVAL this AS ITextSelection2 PTR, BYVAL width_ AS LONG, BYVAL height AS LONG, BYVAL ascent AS LONG, _
      BYVAL nType AS LONG, BYVAL bstrAltText AS BSTR, BYVAL pStream AS IStream PTR) AS HRESULT

END TYPE

TYPE ITextSelection2_
   lpVtbl AS ITextSelection2VTbl PTR
END TYPE

TYPE LPITEXTSELECTION2 AS ITextSelection2 PTR

#endif

' ########################################################################################
' * Interface name: ITextStoryRanges2
' IID: {C241F5E5-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = ITextStoryRanges
' ########################################################################################

#ifndef __ITextStoryRanges2_INTERFACE_DEFINED__
#define __ITextStoryRanges2_INTERFACE_DEFINED__

TYPE ITextStoryRanges2VTbl
   ' // Iispatch interface methods
   QueryInterface AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   ' // Inherited ITextStoryRange interface methods
   _NewEnum AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL ppunkEnum AS IUnknown PTR PTR) AS HRESULT
   Item AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL Index AS LONG, BYVAL ppRange AS ITextRange PTR PTR) AS HRESULT
   GetCount AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   ' // ITextStoryRange2 interface method
   Item2 AS FUNCTION (BYVAL this AS ITextStoryRanges2 PTR, BYVAL Index AS LONG, BYVAL ppRange AS ITextRange2 PTR PTR) AS HRESULT
END TYPE

TYPE ITextStoryRanges2_
   lpVtbl AS ITextStoryRanges2VTbl PTR
END TYPE

TYPE LPITEXTSTORYRANGES2 AS ITextStoryRanges2 PTR

#endif

' ########################################################################################
' * Interface name: ITextStrings
' IID: {C241F5E7-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __ITextStrings_INTERFACE_DEFINED__
#define __ITextStrings_INTERFACE_DEFINED__

TYPE ITextStringsVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextStrings PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextStrings PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   Item AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL Index AS LONG, BYVAL ppRange AS ITextRange2 PTR) AS HRESULT
   GetCount AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL pCount AS LONG PTR) AS HRESULT
   Add AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL bstr AS BSTR) AS HRESULT
   Append AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL pRange AS ITextRange2, BYVAL iString AS LONG) AS HRESULT
   Cat2 AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL iString AS LONG) AS HRESULT
   CatTop2 AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL bstr AS BSTR) AS HRESULT
   DeleteRange AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   EncodeFunction AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL nType AS Long,  BYVAL Align AS LONG, _
       BYVAL Char AS LONG, BYVAL Char1 AS LONG, BYVAL Char2 AS LONG, BYVAL Count AS LONG, _
       BYVAL TeXStyle AS LONG, BYVAL cCol AS LONG, BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   GetCch AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL iString AS LONG, BYVAL pcch AS LONG PTR) AS HRESULT
   InsertNullStr AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL iString AS LONG) AS HRESULT
   MoveBoundary AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL iString AS LONG, BYVAL cch AS LONG) AS HRESULT
   PrefixTop AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL bstr AS BSTR) AS HRESULT
   Remove AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL iString AS LONG, BYVAL cString AS LONG) AS HRESULT
   SetFormattedText AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL pRangeD AS ITextRange2 PTR, BYVAL pRangeS AS ITextRange2 PTR) AS HRESULT
   SetOpCp AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL iString AS LONG, BYVAL cp AS LONG) AS HRESULT
   SuffixTop AS FUNCTION (BYVAL this AS ITextStrings PTR, BYVAL bstr AS BSTR, BYVAL pRange AS ITextRange2 PTR) AS HRESULT
   Swap AS FUNCTION (BYVAL this AS ITextStrings PTR) AS HRESULT
END TYPE

TYPE ITextStrings_
   lpVtbl AS ITextStringsVTbl PTR
END TYPE

TYPE LPITEXTSTRINGS AS ITextStrings PTR

#endif

' ########################################################################################
' * Interface name: ITextRow 
' IID: {C241F5EF-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __ITextRow_INTERFACE_DEFINED__
#define __ITextRow_INTERFACE_DEFINED__

TYPE ITextRowVTbl
   QueryInterface AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextRow PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextRow PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   GetAlignment AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetAlignment AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellCount AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellCount AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellCountCache AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellCountCache AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellIndex AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellIndex AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellMargin AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellMargin AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetHeight AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetHeight AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetIndent AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetIndent AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetKeepTogether AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetKeepTogether AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetKeepWithNext AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetKeepWithNext AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetNestLevel AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetRTL AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetRTL AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellAlignment AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellAlignment AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellColorBack AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellColorBack AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellColorFore AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellColorFore AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellMergeFlags AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellMergeFlags AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellShading AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellShading AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellVerticalText AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellVerticalText AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellWidth AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   SetCellWidth AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   GetCellBorderColors AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pcrLeft AS LONG PTR, _
      BYVAL pcrTop AS LONG PTR, BYVAL pcrRight AS LONG PTR, BYVAL pcrBottom AS LONG PTR) AS HRESULT
   GetCellBorderWidths AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pduLeft AS LONG PTR, _
      BYVAL pduTop AS LONG PTR, BYVAL pduRight AS LONG PTR, BYVAL pduBottom AS LONG PTR) AS HRESULT
   SetCellBorderColors AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL crLeft AS LONG, _
      BYVAL crTop AS LONG, BYVAL crRight AS LONG, BYVAL crBottom AS LONG) AS HRESULT
   SetCellBorderWidths AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL duLeft AS LONG, _
      BYVAL duTop AS LONG, BYVAL duRight AS LONG, BYVAL duBottom AS LONG) AS HRESULT
   Apply AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL cRow AS LONG, BYVAL Flags AS LONG) AS HRESULT
   CanChange AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pValue AS LONG PTR) AS HRESULT
   GetProperty AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL nType AS LONG, BYVAL pValue AS LONG PTR) AS HRESULT
   Insert AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL cRow AS LONG PTR) AS HRESULT
   IsEqual AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pRow AS ITextRow PTR, BYVAL pB AS LONG PTR) AS HRESULT
   Reset AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL Value AS LONG) AS HRESULT
   SetProperty AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL nType AS LONG, BYVAL Value AS LONG) AS HRESULT
END TYPE

TYPE ITextRow_
   lpVtbl AS ITextRowVTbl PTR
END TYPE

TYPE LPITEXTROW AS ITextRow PTR

#endif

' ########################################################################################
' * Interface name: ITextDisplays  
' IID: {C241F5F2-7206-11D8-A2C7-00A0D1D6C6B3}
' Inherited interface = IDispatch
' This interface is currently undefined.
' ########################################################################################

#ifndef __ITextDisplays_INTERFACE_DEFINED__
#define __ITextDisplays_INTERFACE_DEFINED__

TYPE ITextDisplaysVTbl
   ' // IDispatch interface methods
   QueryInterface AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT
   AddRef AS FUNCTION (BYVAL this AS ITextRow PTR) AS ULONG
   Release AS FUNCTION (BYVAL this AS ITextRow PTR) AS ULONG
   GetTypeInfoCount AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT
   GetTypeInfo AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   GetIDsOfNames AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   Invoke AS FUNCTION (BYVAL this AS ITextRow PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE

TYPE ITextDisplays_
   lpVtbl AS ITextDisplaysVTbl PTR
END TYPE

TYPE LPITEXTDISPLAYS AS ITextDisplays PTR

#endif

' ########################################################################################

END NAMESPACE
