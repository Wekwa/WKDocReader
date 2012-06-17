/*
 WKDocReader.m
 WKDocReader

 Copyright 2012 Wyatt Kaufman
 
 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at
 
 http://www.apache.org/licenses/LICENSE-2.0
 
 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 */


#import "WKDocReader.h"
#import "WKDataStream.h"
#import <CoreText/CoreText.h>

#define ARC __has_feature(objc_arc)

#if ARC

#define SAFE_RELEASE(x)
#define SAFE_RETAIN(x)
#define SAFE_AUTORELEASE(x)

#else

#define SAFE_RELEASE(x) [(x) release]
#define SAFE_RETAIN(x) [(x) retain]
#define SAFE_AUTORELEASE(x) [(x) autorelease]

#endif

//A simplified way of returning out of a function if something goes wrong.

#define ENSURE(condition, reasonStr, shouldStopBool) if(!(condition)) { [self setFailureReason:(reasonStr)]; if(shouldStopBool) return; }

enum sectionAttributes {
	sprmSCColumns = 0x500B,
	sprmSBOrientation = 0x301D,
	sprmSXaPage = 0xB01F,
	sprmSYaPage = 0xB020,
	sprmSDxaLeft = 0xB021,
	sprmSDxaRight = 0xB022,
	sprmSDyaTop = 0x9023,
	sprmSDyaBottom = 0x9024,
};

enum paragraphAttributes {
	sprmPJc80 = 0x2403,
	sprmPDxaRight80 = 0x840E,
	sprmPDxaLeft80 = 0x840F,
	sprmPDxaLeft180 = 0x8411,
	sprmPDyaLine = 0x6412,
	sprmPDyaBefore = 0xA413,
	sprmPFBiDi = 0x2441,
};

enum characterAttributes {
	sprmCHighlight = 0x2A0C,
	sprmCFBold = 0x0835,
	sprmCFItalic = 0x0836,
	sprmCFOutline = 0x0838,
	sprmCKul = 0x2A3E,
	sprmCIco = 0x2A42,
	sprmCHps = 0x4A43,
	sprmCIss = 0x2A48,
	sprmCRgFtc0 = 0x4A4F,
	sprmCRgFtc1 = 0x4A50,
	sprmCRgFtc2 = 0x4A51,
	sprmCCv = 0x6870,
	sprmCCvUl = 0x6877,
};

NSString *const WKColumnCountAttributeName = @"WKColumnCountAttributeName";
NSString *const WKPageOrientationAttributeName = @"WKPageOrientationAttributeName";
NSString *const WKPageWidthAttributeName = @"WKPageWidthAttributeName";
NSString *const WKPageHeightAttributeName = @"WKPageHeightAttributeName";
NSString *const WKLeftMarginAttributeName = @"WKLeftMarginAttributeName";
NSString *const WKRightMarginAttributeName = @"WKRightMarginAttributeName";
NSString *const WKTopMarginAttributeName = @"WKTopMarginAttributeName";
NSString *const WKBottomMarginAttributeName = @"WKBottomMarginAttributeName";
NSString *const WKBackgroundColorAttributeName = @"WKBackgroundColorAttributeName";

NSString *const WKReadOnlyDocumentAttribute = @"WKReadOnlyDocumentAttribute";
NSString *const WKHideSpellingErrorsDocumentAttribute = @"WKHideSpellingErrorsDocumentAttribute";
NSString *const WKHideGrammarErrorsDocumentAttribute  = @"WKHideGrammarErrorsDocumentAttribute";
NSString *const WKDefaultTabIntervalDocumentAttribute = @"WKDefaultTabIntervalDocumentAttribute";
NSString *const WKCreationTimeDocumentAttribute = @"WKCreationTimeDocumentAttribute";
NSString *const WKModificationTimeDocumentAttribute = @"WKModificationTimeDocumentAttribute"; 
NSString *const WKViewModeDocumentAttribute = @"WKViewModeDocumentAttribute";
NSString *const WKViewZoomDocumentAttribute = @"WKViewZoomDocumentAttribute";
NSString *const WKAutosizeDocumentAttribute = @"WKAutosizeDocumentAttribute";

#pragma mark -
#pragma mark NSMutableAttributedString conveniences

@interface NSMutableAttributedString (SequentialAttributes)

-(void)fillStringWithDefaultAttributes;
-(void)setAttribute:(CFStringRef)attribute toValue:(id)value startingAtIndex:(NSUInteger)index endingAtIndex:(NSUInteger)end;

@end

@implementation NSMutableAttributedString (SequentialAttributes)

-(void)fillStringWithDefaultAttributes
{
	
	/*NSAttributedStrings come empty by default. In Word files, all attributes
	 are based on previously-existing attributes, so we have to put in some default
	 for these fields.
	 */
	CTFontRef defaultFont = CTFontCreateWithName(CFSTR("Helvetica"), 12.0, NULL);
	[self setAttribute:kCTFontAttributeName toValue:(id)defaultFont startingAtIndex:0 endingAtIndex:[self length]];
	CFRelease(defaultFont);
	
	CTTextAlignment alignment = kCTLeftTextAlignment;
	CGFloat rightIndent = 0.0;
	CGFloat leftIndent = 0.0;
	CGFloat firstLeftIndent = 0.0;
	CGFloat lineHeightMultiple = 1.0;
	CGFloat minLineHeight = 0.0;
	CGFloat maxLineHeight = 0.0;
	CGFloat paragraphSpacingBefore = 0.0;
	CTWritingDirection writingDirection = kCTWritingDirectionNatural;
	CTLineBreakMode breakMode = kCTLineBreakByCharWrapping;
	
	CTParagraphStyleSetting newParagraphSettings[] = {
		{ kCTParagraphStyleSpecifierAlignment, sizeof(alignment), &alignment },
		{ kCTParagraphStyleSpecifierTailIndent, sizeof(rightIndent), &rightIndent },
		{ kCTParagraphStyleSpecifierHeadIndent, sizeof(leftIndent), &leftIndent },
		{ kCTParagraphStyleSpecifierFirstLineHeadIndent, sizeof(firstLeftIndent), &firstLeftIndent },
		{ kCTParagraphStyleSpecifierLineHeightMultiple, sizeof(lineHeightMultiple), &lineHeightMultiple },
		{ kCTParagraphStyleSpecifierMinimumLineHeight, sizeof(minLineHeight), &minLineHeight },
		{ kCTParagraphStyleSpecifierMaximumLineHeight, sizeof(maxLineHeight), &maxLineHeight },
		{ kCTParagraphStyleSpecifierParagraphSpacingBefore, sizeof(paragraphSpacingBefore), &paragraphSpacingBefore },
		{ kCTParagraphStyleSpecifierBaseWritingDirection, sizeof(writingDirection), &writingDirection },
		{ kCTParagraphStyleSpecifierLineBreakMode, sizeof(CTLineBreakMode), &breakMode }
	};
	CTParagraphStyleRef newParagraphStyle = CTParagraphStyleCreate(newParagraphSettings, 10);
	[self setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParagraphStyle startingAtIndex:0 endingAtIndex:[self length]];
	CFRelease(newParagraphStyle);
	
}

-(void)setAttribute:(CFStringRef)attribute toValue:(id)value startingAtIndex:(NSUInteger)index endingAtIndex:(NSUInteger)endIndex;
{
	//Won't leak memory like -addAttribute:value:range does
	if(endIndex > [self length]) return;
	
	NSMutableDictionary *existingAttributes = [[self attributesAtIndex:index effectiveRange:NULL] mutableCopy];
	if([existingAttributes objectForKey:(NSString *)attribute]) {
		[existingAttributes removeObjectForKey:(NSString *)attribute];
	}
	[existingAttributes setObject:value forKey:(NSString *)attribute];
	[self setAttributes:existingAttributes range:NSMakeRange(index, endIndex - index)];
	SAFE_RELEASE(existingAttributes);
}

@end

#pragma mark -
#pragma mark Internal method declarations

@interface WKDocReader (Private)

-(void)read;
-(void)readHeaderAndAssembleStreams;
-(void)readFATFromSector:(NSUInteger)sector;
-(void)readMiniFATFromSector:(NSUInteger)sector;
-(void)readFileInformationBlock;
-(void)readTextContent;
-(void)readSectionAttributes;
-(void)readFontTable;
-(void)readStylesheet;
-(void)readFormatting;
-(void)readFKPWithData:(NSData *)fkpData isPapxFKP:(BOOL)isPapx;
-(void)readDocumentProperties;

-(CTParagraphStyleRef)createModifiedParagraphStyle:(CTParagraphStyleRef)originalStyle specifier:(CTParagraphStyleSpecifier)specifier newValue:(void *)valuePtr;
-(uint32_t)convertFCtoCP:(uint32_t)fc;
-(uint16_t)sizeOfOperandForSprm:(uint16_t)sprm;
-(CGColorRef)CGColorWithICO:(uint8_t)ico;
-(CGColorRef)CGColorWithCCv:(uint32_t)ccv;
-(NSDate *)dateWithDTTM:(uint32_t)dttm;

-(void)applySepx:(NSData *)sepx fromIndex:(uint32_t)start toIndex:(uint32_t)index;
-(void)applyFormatting:(NSData *)formatting fromIndex:(uint32_t)start toIndex:(uint32_t)index;

-(void)setFailureReason:(NSString *)reason;

@end

@implementation WKDocReader {
	
	WKDataStream *_fileStream;
	WKDataStream *_wordStream;
	WKDataStream *_tableStream;
	
	
	//These members are named to correspond with MS-CFB and MS-DOC documentation.
	struct {
		uint16_t pageSize;
		uint32_t numberOfFATSectors;
		uint32_t firstDirectorySectorLocation;
		uint32_t firstMiniFATSectorLocation;
		uint32_t numberOfMiniFATSectors;
		uint32_t firstMiniDIFATSectorLocation;
		uint32_t numberOfDIFATSectors;
		uint32_t firstSectorOfMinistream;
		uint32_t ministreamLength;
		BOOL hasPictures;
		
		uint32_t ccpText; //The length of the text.
		uint32_t fcStshf; //The location in the table stream of the stylesheet
		uint32_t lcbStshf; //The size in bytes of the stylesheet
		uint32_t fcPlcfSed; //The location in the table stream of a SED. (see -readSectionAttributes for more info.)
		uint32_t lcbPlcfSed; //The size of the SED
		uint32_t fcPlcfBteChpx; //The location in the table stream of a plex for character formatting. (see -readFormatting for more info.)
		uint32_t lcbPlcfBteChpx; //The size in bytes of the plex for character formatting
		uint32_t fcPlcfBtePapx; //The location in the table stream of a plex for paragraph formatting. (see -readFormatting for more info.)
		uint32_t lcbPlcfBtePapx; //The size in bytes of the plex for paragraph formatting.
		uint32_t fcSttbfFfn; //The location in the table stream of the font table.
		uint32_t lcbSttbfFfn; //The size in bytes of the font table.
		uint32_t fcDop; //The location in the table stream of the document properties.
		uint32_t lcbDop; //The size in bytes of the document properties.
		uint32_t fcClx; //The location in the table stream of a Clx. (see -readTextContent for more info.)
		uint32_t lcbClx; //The size in bytes of the Clx.
		
		uint32_t textStartLocation; //The location in the word stream where the text starts.
		BOOL fcCompressed; //If YES, the text in the document is stored in UTF-8. Otherwise, it is stored as UTF-16.
	} _docInfo;
	
	NSMutableAttributedString *_attributedString;
	NSMutableArray *_fileAllocationTables;
	NSMutableArray *_miniFileAllocationTables;
	NSMutableArray *_fontNames;
	NSMutableDictionary *_styleGrpprlDatas; //Object = data for style. Key = NSNumber of istd
	NSMutableDictionary *_documentAttributes;
	NSError *_error;
}

#pragma mark -
#pragma mark Creation/Deletion

-(id)initWithDocFormatData:(NSData *)data
{
	self = [super init];
	if(self) {
		_fileStream = [[WKDataStream alloc] initWithData:data];
		_fontNames = [[NSMutableArray alloc] init];
		_fileAllocationTables = [[NSMutableArray alloc] init];
		_miniFileAllocationTables = [[NSMutableArray alloc] init];
		_styleGrpprlDatas = [[NSMutableDictionary alloc] init];
		_documentAttributes = [[NSMutableDictionary alloc] init];
		[self read];
	}
	return self;
}

-(id)initWithContentsOfFile:(NSString *)filePath
{
	NSData *data = [NSData dataWithContentsOfFile:filePath];
	return [self initWithDocFormatData:data];
}

#if !ARC

-(void)dealloc
{
	[_attributedString release];
	[_documentAttributes release];
	[_fileStream release];
	[_tableStream release];
	[_wordStream release];
	[_fileAllocationTables release];
	[_miniFileAllocationTables release];
	[_fontNames release];
	[_styleGrpprlDatas release];
	[_error release];
	[super dealloc];
}

#endif

#pragma mark -
#pragma mark Reading

-(void)readHeaderAndAssembleStreams
{
	/*In a Word document, (OLE binary), the data is not necessarily meant to be read in the order
	 that it's written to the file. It can be thought of more like a file system, where objects
	 are stored in a hierarchy, rather than in sequence. 
	 
	 However, the data has to be read in *some* order, so the header (the first 512 bytes of the document)
	 will tell us where to start.
	 */
	
	uint32_t headerSignature1 = [_fileStream readUInt32];
	uint32_t headerSignature2 = [_fileStream readUInt32];
	ENSURE((headerSignature1 == 0xE011CFD0 && headerSignature2 == 0xE11AB1A1), @"Invalid header signature -- file is not a .doc file.", YES);
	
	[_fileStream skipBytes:22];
	uint16_t sectorShift = [_fileStream readUInt16];
	_docInfo.pageSize = 512;
	if(sectorShift == 0x000C) _docInfo.pageSize = 4096;
	[_fileStream skipBytes:12];
	
	_docInfo.numberOfFATSectors = [_fileStream readUInt32];
	_docInfo.firstDirectorySectorLocation = [_fileStream readUInt32];
	[_fileStream skipBytes:8];
	_docInfo.firstMiniFATSectorLocation = [_fileStream readUInt32];
	_docInfo.numberOfMiniFATSectors = [_fileStream readUInt32];
	_docInfo.firstMiniDIFATSectorLocation = [_fileStream readUInt32];
	_docInfo.numberOfDIFATSectors = [_fileStream readUInt32];
	
	/*Word documents are separated into 512-byte chunks called "sectors", or "pages."
	 The File Allocation Tables (FATs) describe how to properly arrange the sectors
	 in the document to be read.
	
	 The DIFAT, the final 436 bytes of the header, gives the sector numbers of the FATs.
	 */
	
	while([_fileStream index] < _docInfo.pageSize) {
		uint32_t fatIndex = [_fileStream readUInt32];
		if(fatIndex != 0xFFFFFFFF)
			[self readFATFromSector:fatIndex];
	};
	
	/*The document MAY also contain Mini-FAT sectors. These are for streams that are
	 less than 4096 bytes long.
	 */
	
	if(_docInfo.numberOfMiniFATSectors > 0) {
		for(int i = 0; i < _docInfo.numberOfMiniFATSectors; i++) {
			[self readMiniFATFromSector:(_docInfo.firstMiniFATSectorLocation + i)];
		}
	}
	
	/*The Table stream may be described as either 0Table or 1Table. Here we peek ahead
	 to the FIB to find out which on it is. (See -readFileInformationBlock for more info.)*/
	int tableNumber;
	[_fileStream skipToByte:522];
	tableNumber = ((([_fileStream readUInt16]) & 0x0200) == 0x0200);
	
	[_fileStream skipToByte:(_docInfo.firstDirectorySectorLocation + 1) * _docInfo.pageSize];
	
	/*The Directory is a list of all the streams in the document. An entry for a stream is
	 128 bytes long and describes, among other things, where the stream starts and how long it is. 
	 It first lists the Root Entry, which acts as a storage container for all streams. Following 
	 that are all the streams of the document. There may be several, but the only ones needed
	 for this purpose are the WordDocument stream and the Table Stream.
	 */
	
	[_fileStream skipBytes:64];
	uint16_t rootEntryLength = [_fileStream readUInt16];
	[_fileStream skipToByte:(_docInfo.firstDirectorySectorLocation + 1) * _docInfo.pageSize];
	uint8_t *rootEntryNameBytes = [_fileStream readBytes:(rootEntryLength - 2)];
	NSString *rootEntryName = [[NSString alloc] initWithBytes:rootEntryNameBytes length:(rootEntryLength - 2) encoding:NSUTF16LittleEndianStringEncoding];
	ENSURE([rootEntryName isEqualToString:@"Root Entry"], @"No Root Entry in Directory", YES);
	SAFE_RELEASE(rootEntryName);
	[_fileStream skipBytes:64 - rootEntryLength + 54];
	_docInfo.firstSectorOfMinistream = [_fileStream readUInt32];
	_docInfo.ministreamLength = [_fileStream readUInt32];
	[_fileStream skipBytes:4];
	
	BOOL hasWordStream = NO, hasTableStream = NO;
	while((!hasWordStream || !hasTableStream) && ![_fileStream isAtEnd]) {
		[_fileStream skipBytes:64];
		uint16_t streamNameLength = [_fileStream readUInt16];
		[_fileStream skipBytes:-66];
		uint8_t *streamNameBytes = [_fileStream readBytes:(streamNameLength - 2)];
		NSString *streamName = [[NSString alloc] initWithBytes:streamNameBytes length:(streamNameLength - 2) encoding:NSUTF16LittleEndianStringEncoding];
		[_fileStream skipBytes:64 - streamNameLength + 54];
		uint32_t sector = [_fileStream readUInt32];
		if((sector + 1) * (_docInfo.pageSize) > [[_fileStream data] length]) continue;
		NSMutableData *streamData = [NSMutableData data];
		uint32_t streamLength = [_fileStream readUInt32];
		if(streamLength < 4096) {
			NSData *miniData = [[_fileStream data] subdataWithRange:NSMakeRange(((_docInfo.firstSectorOfMinistream + 1) * _docInfo.pageSize), _docInfo.ministreamLength)];
			while(sector != 0xFFFFFFFE) {
				NSData *sectorData = [miniData subdataWithRange:NSMakeRange((sector) * 64, 64)];
				[streamData appendData:sectorData];
				sector = [[_miniFileAllocationTables objectAtIndex:sector] unsignedIntValue];
			}		
		} else {
			while(sector != 0xFFFFFFFE) {
				NSData *sectorData = [[_fileStream data] subdataWithRange:NSMakeRange((sector + 1) * _docInfo.pageSize, _docInfo.pageSize)];
				[streamData appendData:sectorData];
				sector = [[_fileAllocationTables objectAtIndex:sector] unsignedIntValue];
			}
		}
		if([streamName isEqualToString:@"WordDocument"]) {
			_wordStream = [[WKDataStream alloc] initWithData:streamData];
			hasWordStream = YES;
		} else if([streamName isEqualToString:[NSString stringWithFormat:@"%iTable", tableNumber]]) {
			_tableStream = [[WKDataStream alloc] initWithData:streamData];
			hasTableStream = YES;
		}
		
		SAFE_RELEASE(streamName);
		[_fileStream skipBytes:4];
	}
	
	
	[_fileStream skipToByte:(_docInfo.firstDirectorySectorLocation + 1) * _docInfo.pageSize];
	
}

-(void)readFATFromSector:(NSUInteger)sector
{
	/*A FAT is one sector consisting of an array of 4-byte integers. Each integer in the
	 chain refers to the sector number that comes next. If the integer is 0xFFFFFFFE,
	 it signifies the end of a stream.
	 */
	
	[_fileStream skipToByte:((sector + 1) * _docInfo.pageSize)];
	while(![_fileStream isAtEnd]) {
		uint32_t chain = [_fileStream readUInt32];
		[_fileAllocationTables addObject:[NSNumber numberWithUnsignedInt:chain]];
	}
}

-(void)readMiniFATFromSector:(NSUInteger)sector
{
	[_fileStream skipToByte:((sector + 1) * _docInfo.pageSize)];
	while(![_fileStream isAtEnd]) {
		uint32_t chain = [_fileStream readUInt32];
		[_miniFileAllocationTables addObject:[NSNumber numberWithUnsignedInt:chain]];
	}
}

-(void)readFileInformationBlock
{
	/*The FileInformationBlock is a structure at the beginning of the Word stream. It
	 describes where other objects are located in the file, as well as some miscellaneous
	 verification info.
	 */
	//FibBase
	[_wordStream skipToByte:0];
	uint16_t wIdent = [_wordStream readUInt16];
	ENSURE((wIdent == 0xA5EC), @"Document is not a Word Binary File.", YES);
	[_wordStream skipBytes:8];
	uint16_t flags16 = [_wordStream readUInt16];
	_docInfo.hasPictures = ((flags16 & 0x0008) == 0x0008);
	ENSURE(((flags16 & 0x0100) != 0x0100), @"Document is encrypted. (Encrypted document unsupported.)", YES);
	BOOL readOnly = ((flags16 & 0x0400) == 0x0400);
	[_documentAttributes setObject:[NSNumber numberWithBool:readOnly] forKey:WKReadOnlyDocumentAttribute];
	ENSURE(((flags16 & 0x8000) != 0x8000), @"Document is obfuscated. (XOR obfuscation unsupported.)", YES);
	
	//FibRgW97
	[_wordStream skipBytes:20];
	ENSURE(([_wordStream readUInt16] == 0x000E), @"FibRgW97 is the wrong length.", YES);
	[_wordStream skipBytes:28];
	
	//FibRgLw97
	ENSURE(([_wordStream readUInt16] == 0x0016), @"FibRgLw97 is the wrong length.", YES);
	[_wordStream skipBytes:12];
	_docInfo.ccpText = [_wordStream readUInt32];
	[_wordStream skipBytes:72];
	
	//FibRgFcLcbBlob
	[_wordStream skipBytes:10];
	_docInfo.fcStshf = [_wordStream readUInt32];
	_docInfo.lcbStshf = [_wordStream readUInt32];
	[_wordStream skipBytes:32];
	_docInfo.fcPlcfSed = [_wordStream readUInt32];
	_docInfo.lcbPlcfSed = [_wordStream readUInt32];
	[_wordStream skipBytes:40];
	_docInfo.fcPlcfBteChpx = [_wordStream readUInt32];
	_docInfo.lcbPlcfBteChpx = [_wordStream readUInt32];
	_docInfo.fcPlcfBtePapx = [_wordStream readUInt32];
	_docInfo.lcbPlcfBtePapx = [_wordStream readUInt32];
	[_wordStream skipBytes:8];
	_docInfo.fcSttbfFfn = [_wordStream readUInt32];
	_docInfo.lcbSttbfFfn = [_wordStream readUInt32];
	[_wordStream skipBytes:120];
	_docInfo.fcDop = [_wordStream readUInt32];
	_docInfo.lcbDop = [_wordStream readUInt32];
	[_wordStream skipBytes:8];
	_docInfo.fcClx = [_wordStream readUInt32];
	_docInfo.lcbClx = [_wordStream readUInt32];
}

-(void)readTextContent
{
	/*In a Word document, all text for *every* part of the document is stored
	 in the same area. ie, the text for the header and footer are written 
	 *immediately* after the main body text. The Clx structure in the table stream
	 defines what ranges of the text belong to which parts of the document.
	 (This parser will only use the range designated for the main body text.)
	 */
	[_tableStream skipToByte:_docInfo.fcClx];
	uint8_t clxt = [_tableStream peekUInt8];
	while(clxt != 0x02) {
		[_tableStream skipBytes:1];
		uint16_t cbGrpprl = [_tableStream readUInt16];
		[_tableStream skipBytes:cbGrpprl];
		clxt = [_tableStream peekUInt8];
	}
	[_tableStream skipBytes:1];
	uint32_t lcbPlcPcd = [_tableStream readUInt32];
	uint32_t pcdCount = (lcbPlcPcd - 4) / 12;
	[_tableStream skipBytes:((pcdCount + 1) * 4)];
	[_tableStream skipBytes:2];
	uint32_t fc = [_tableStream readUInt32];
	 _docInfo.textStartLocation = (fc & 0x3FFFFFFF);
	_docInfo.fcCompressed = ((fc & 0x40000000) == 0x40000000);
	
	if(_docInfo.fcCompressed) {
		_docInfo.textStartLocation /= 2;
	}
	[_wordStream skipToByte:_docInfo.textStartLocation];
	uint32_t realTextLength = (_docInfo.ccpText * (_docInfo.fcCompressed ? 1 : 2));
	uint8_t *textBytes = [_wordStream readBytes:realTextLength];
	NSString *textContent = SAFE_AUTORELEASE([[NSString alloc] initWithBytes:textBytes length:realTextLength encoding:(_docInfo.fcCompressed ? NSUTF8StringEncoding : NSUTF16LittleEndianStringEncoding)]);
	_attributedString = [[NSMutableAttributedString alloc] initWithString:textContent];
	[_attributedString fillStringWithDefaultAttributes];
}

-(void)applySepx:(NSData *)sepx fromIndex:(uint32_t)start toIndex:(uint32_t)index
{
	/*Sepxs define the properties for a section. (This includes columns, page size, etc.) 
	 It is analogous to Chpxs and Papx for character or paragraph formatting. See
	 -applyFormatting:fromIndex:toIndex for more information about how this data is stored.
	 */
	
	WKDataStream *sepxStream = [[WKDataStream alloc] initWithData:sepx];
	while(![sepxStream isAtEnd]) {
		uint16_t sprm = [sepxStream readUInt16];
		switch (sprm) {
				
			case sprmSCColumns: {
				uint16_t columnCount = [sepxStream readUInt16];
				[_attributedString setAttribute:(CFStringRef)WKColumnCountAttributeName toValue:[NSNumber numberWithShort:(columnCount + 1)] startingAtIndex:start endingAtIndex:index];
				 break;
			}
				
			case sprmSBOrientation: {
				uint8_t orientation = [sepxStream readUInt8];
				[_attributedString setAttribute:(CFStringRef)WKPageOrientationAttributeName toValue:[NSNumber numberWithShort:(orientation - 1)] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmSXaPage: {
				uint16_t pageWidth = ([sepxStream readUInt16] / 20);
				[_attributedString setAttribute:(CFStringRef)WKPageWidthAttributeName toValue:[NSNumber numberWithShort:pageWidth] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmSYaPage: {
				uint16_t pageWidth = ([sepxStream readUInt16] / 20);
				[_attributedString setAttribute:(CFStringRef)WKPageHeightAttributeName toValue:[NSNumber numberWithShort:pageWidth] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmSDxaLeft: {
				uint16_t marginWidth = [sepxStream readUInt16];
				[_attributedString setAttribute:(CFStringRef)WKLeftMarginAttributeName toValue:[NSNumber numberWithShort:marginWidth] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmSDxaRight: {
				uint16_t marginWidth = [sepxStream readUInt16];
				[_attributedString setAttribute:(CFStringRef)WKRightMarginAttributeName toValue:[NSNumber numberWithShort:marginWidth] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmSDyaTop: {
				uint16_t marginHeight = [sepxStream readUInt16];
				[_attributedString setAttribute:(CFStringRef)WKTopMarginAttributeName toValue:[NSNumber numberWithShort:marginHeight] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmSDyaBottom: {
				uint16_t marginHeight = [sepxStream readUInt16];
				[_attributedString setAttribute:(CFStringRef)WKBottomMarginAttributeName toValue:[NSNumber numberWithShort:marginHeight] startingAtIndex:start endingAtIndex:index];
				break;
			}
			
			default:
				[sepxStream skipBytes:[self sizeOfOperandForSprm:sprm]];
				break;
		}
	}
	SAFE_RELEASE(sepxStream);
}
;
-(void)applyFormatting:(NSData *)formatting fromIndex:(uint32_t)start toIndex:(uint32_t)index
{
	/*Formatting data is known as Chpx for character formatting or Papx for paragraph formatting.
	 A Chpx consists of an array of SPRM/operand pairs. A SPRM (Single Property Modifier) is a
	 16-bit integer that indicates what property is being modified. (ie, bold, italic, color, etc.)
	 The operand that follows it varies in size (see -sizeOfOperandForSprm:), and describes the
	 argument for the sprm's action. ie, if the sprm indicated a change in font size, the operand
	 would be the new font size.
	 */
	WKDataStream *fmtStream = [[WKDataStream alloc] initWithData:formatting];
	if([formatting length] < 3) return;
	if(start + 1 > [_attributedString length]) return;
	if(index + 1 > [_attributedString length]) {
		index = [_attributedString length] - 1;
	}
	while(![fmtStream isAtEnd]) {
		
		uint16_t sprm = [fmtStream readUInt16];
		switch (sprm) {
			case sprmPJc80: {
				CTParagraphStyleRef existingParaStyle = (CTParagraphStyleRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTParagraphStyleAttributeName];
				uint8_t alignmentInt = [fmtStream readUInt8];
				CTTextAlignment alignment;
				switch(alignmentInt) {
					case 0x00:
						alignment = kCTLeftTextAlignment;
						break;
					case 0x01:
						alignment = kCTCenterTextAlignment;
						break;
					case 0x02:
						alignment = kCTRightTextAlignment;
						break;
					case 0x03:
					case 0x04:
					case 0x05:
						alignment = kCTJustifiedTextAlignment;
						break;
				}
				CTParagraphStyleRef newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierAlignment newValue:&alignment];
				[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
				CFRelease(newParaStyle);
				break;
			}
			
			case sprmPDxaRight80: {
				
				CTParagraphStyleRef existingParaStyle = (CTParagraphStyleRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTParagraphStyleAttributeName];
				uint16_t xas = [fmtStream readUInt16];
				CGFloat rightIndent = ((CGFloat)xas / 20.0);

				CTParagraphStyleRef newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierTailIndent newValue:&rightIndent];
				[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
				CFRelease(newParaStyle);
				break;
			}
				
			case sprmPDxaLeft80: {
				
				CTParagraphStyleRef existingParaStyle = (CTParagraphStyleRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTParagraphStyleAttributeName];
				uint16_t xas = [fmtStream readUInt16];
				CGFloat leftIndent = ((CGFloat)xas / 20.0);
				
				CTParagraphStyleRef newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierHeadIndent newValue:&leftIndent];
				[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
				CFRelease(newParaStyle);
				break;
			}
				
			case sprmPDxaLeft180: {
				
				CTParagraphStyleRef existingParaStyle = (CTParagraphStyleRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTParagraphStyleAttributeName];
				uint16_t xas = [fmtStream readUInt16];
				CGFloat firstLeftIndent = ((CGFloat)xas / 20.0);
				CTParagraphStyleRef newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierFirstLineHeadIndent newValue:&firstLeftIndent];
				[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
				CFRelease(newParaStyle);
				break;
			}
				
			case sprmPDyaLine: {
				CTParagraphStyleRef existingParaStyle = (CTParagraphStyleRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTParagraphStyleAttributeName];
				uint16_t dyaLine = [fmtStream readUInt16];
				uint16_t fMultiLinespace = [fmtStream readUInt16];
				CGFloat minLineHeight, maxLineHeight, lineHeightMultiple;
				CTParagraphStyleRef newParaStyle;
				if(dyaLine >= 0x8440) {
					minLineHeight = maxLineHeight = (0x10000 - dyaLine);
					newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierMinimumLineHeight newValue:&minLineHeight];
					newParaStyle = [self createModifiedParagraphStyle:newParaStyle specifier:kCTParagraphStyleSpecifierMaximumLineHeight newValue:&maxLineHeight];
					[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
					CFRelease(newParaStyle);
					break;
				}
				if(fMultiLinespace == 0x0001) {
					lineHeightMultiple = ((CGFloat)dyaLine / 240.0);
					newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierLineHeightMultiple newValue:&lineHeightMultiple];
					[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
					CFRelease(newParaStyle);
					break;
				}
				minLineHeight = maxLineHeight = ((CGFloat)dyaLine / 20.0);
				newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierMinimumLineHeight newValue:&minLineHeight];
				newParaStyle = [self createModifiedParagraphStyle:newParaStyle specifier:kCTParagraphStyleSpecifierMaximumLineHeight newValue:&maxLineHeight];
				[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
				CFRelease(newParaStyle);
				break;
			}
				
			case sprmPDyaBefore: {
				
				CTParagraphStyleRef existingParaStyle = (CTParagraphStyleRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTParagraphStyleAttributeName];
				uint16_t dyaBefore = [fmtStream readUInt16];
				CGFloat paragraphSpacingBefore = ((CGFloat)dyaBefore / 20.0);
				CTParagraphStyleRef newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierParagraphSpacingBefore newValue:&paragraphSpacingBefore];
				[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
				CFRelease(newParaStyle);
				break;
			}
				
			case sprmPFBiDi: {
				CTParagraphStyleRef existingParaStyle = (CTParagraphStyleRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTParagraphStyleAttributeName];
				uint8_t direction = [fmtStream readUInt8];
				CTWritingDirection writingDirection;
				if(direction) {
					writingDirection = kCTWritingDirectionRightToLeft;
					break;
				}
				writingDirection = kCTWritingDirectionNatural;
				CTParagraphStyleRef newParaStyle = [self createModifiedParagraphStyle:existingParaStyle specifier:kCTParagraphStyleSpecifierBaseWritingDirection newValue:&writingDirection];
				[_attributedString setAttribute:kCTParagraphStyleAttributeName toValue:(id)newParaStyle startingAtIndex:start endingAtIndex:index];
				CFRelease(newParaStyle);
				break;
			}
				
			case sprmCHighlight: {
				uint8_t ico = [fmtStream readUInt8];
				CGColorRef color = [self CGColorWithICO:ico];
				[_attributedString setAttribute:(CFStringRef)WKBackgroundColorAttributeName toValue:(id)color startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmCFBold: {
				BOOL shouldBeBold = ([fmtStream readUInt8] == 0x01);
				CTFontRef existingFont = (CTFontRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTFontAttributeName];
				CTFontSymbolicTraits preTraits = CTFontGetSymbolicTraits(existingFont);
				CTFontSymbolicTraits desiredTrait = 0;
				CTFontSymbolicTraits traitMask = kCTFontBoldTrait;
				
				if(shouldBeBold) {
					if(!(preTraits & kCTFontBoldTrait)) {
						desiredTrait = kCTFontBoldTrait;
					}
				} else {
					if(preTraits & kCTFontBoldTrait) {
						desiredTrait = 0;
					} else {
						desiredTrait = kCTFontBoldTrait;
					}
				}

				CTFontRef newFont = CTFontCreateCopyWithSymbolicTraits(existingFont, CTFontGetSize(existingFont), NULL, desiredTrait, traitMask);
				[_attributedString setAttribute:kCTFontAttributeName toValue:(id)newFont startingAtIndex:start endingAtIndex:index];
				
				CFRelease(newFont);
				break;
			}
				
			case sprmCFItalic: {
				BOOL shouldBeItalic = ([fmtStream readUInt8] == 0x01);
				if(shouldBeItalic) {
				}
				CTFontRef existingFont = (CTFontRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTFontAttributeName];
				CTFontSymbolicTraits preTraits = CTFontGetSymbolicTraits(existingFont);
				CTFontSymbolicTraits desiredTrait = 0;
				CTFontSymbolicTraits traitMask = kCTFontItalicTrait;
				
				if(shouldBeItalic) {
					if(!(preTraits & kCTFontItalicTrait)) {
						desiredTrait = kCTFontItalicTrait;
					}
				} else {
					if(preTraits & kCTFontItalicTrait) {
						desiredTrait = 0;
					} else {
						desiredTrait = kCTFontItalicTrait;
					}
				}
				
				CTFontRef newFont = CTFontCreateCopyWithSymbolicTraits(existingFont, CTFontGetSize(existingFont), NULL, desiredTrait, traitMask);
				[_attributedString setAttribute:kCTFontAttributeName toValue:(id)newFont startingAtIndex:start endingAtIndex:index];
				
				CFRelease(newFont);
				break;
			}
				
			case sprmCFOutline: {
				uint8_t isOutline = ([fmtStream readUInt8] == 0x01);
				[_attributedString setAttribute:kCTStrokeWidthAttributeName toValue:[NSNumber numberWithChar:isOutline] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmCKul: {
				uint8_t kul = [fmtStream readUInt8];
				CTUnderlineStyle style = kCTUnderlineStyleNone;
				
				switch(kul) {
					case 0x01: style = kCTUnderlineStyleSingle; break;
					case 0x03: style = kCTUnderlineStyleDouble; break;
					case 0x04: style = kCTUnderlineStyleSingle | kCTUnderlinePatternDot; break;
					case 0x06: style = kCTUnderlineStyleThick; break;
					case 0x07: style = kCTUnderlineStyleSingle | kCTUnderlinePatternDash; break;
					case 0x09: style = kCTUnderlineStyleSingle | kCTUnderlinePatternDashDot; break;
					case 0x0A: style = kCTUnderlineStyleSingle | kCTUnderlinePatternDashDotDot; break;
					case 0x14: style = kCTUnderlineStyleThick | kCTUnderlinePatternDot; break;
					case 0x17: style = kCTUnderlineStyleThick | kCTUnderlinePatternDash; break;
					case 0x19: style = kCTUnderlineStyleThick | kCTUnderlinePatternDashDot; break;
					case 0x1A: style = kCTUnderlineStyleThick | kCTUnderlinePatternDashDotDot; break;
				}
				
				[_attributedString setAttribute:kCTUnderlineStyleAttributeName toValue:[NSNumber numberWithUnsignedInt:style] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmCIco: {
				uint8_t ico = [fmtStream readUInt8];
				CGColorRef textColor = [self CGColorWithICO:ico];
				[_attributedString setAttribute:kCTForegroundColorAttributeName toValue:(id)textColor startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmCHps: {
				uint16_t halfPointSize = [fmtStream readUInt16];
				CGFloat fontSize = ((CGFloat)halfPointSize / 2.0);
				CTFontRef existingFont = (CTFontRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTFontAttributeName];
				CFStringRef fontName = CTFontCopyName(existingFont, kCTFontFullNameKey);
				CTFontRef newFont = CTFontCreateWithName(fontName, fontSize, NULL);
				[_attributedString setAttribute:kCTFontAttributeName toValue:(id)newFont startingAtIndex:start endingAtIndex:index];
				CFRelease(fontName);
				CFRelease(newFont);
				break;
			}
				
			case sprmCIss: {
				uint8_t baseline = [fmtStream readUInt8];
				int ctBaseline = baseline;
				if(ctBaseline == 0x02) ctBaseline = -1;
				[_attributedString setAttribute:kCTSuperscriptAttributeName toValue:[NSNumber numberWithInt:ctBaseline] startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmCRgFtc0:
			case sprmCRgFtc1:
			case sprmCRgFtc2: {
				uint16_t fontIndex = [fmtStream readUInt16];
				CTFontRef existingFont = (CTFontRef)[[_attributedString attributesAtIndex:start effectiveRange:NULL] objectForKey:(id)kCTFontAttributeName];
			//	CTFontSymbolicTraits traits = CTFontGetSymbolicTraits(existingFont);
				CGFloat size = CTFontGetSize(existingFont);
				
				NSString *fontName = [_fontNames objectAtIndex:fontIndex];
				CTFontRef newFontSansTraits = CTFontCreateWithName((CFStringRef)fontName, size, NULL);
				//CTFontRef newFont = CTFontCreateCopyWithSymbolicTraits(newFontSansTraits, size, NULL, traits, traits);
				//CFRelease(newFontSansTraits);
				[_attributedString setAttribute:kCTFontAttributeName toValue:(id)newFontSansTraits startingAtIndex:start endingAtIndex:index];
				
				break;
			}
				
			case sprmCCv: {
				uint32_t ccv = [fmtStream readUInt32];
				CGColorRef textColor = [self CGColorWithCCv:ccv];
				[_attributedString setAttribute:kCTForegroundColorAttributeName toValue:(id)textColor startingAtIndex:start endingAtIndex:index];
				break;
			}
				
			case sprmCCvUl: {
				uint32_t ccv = [fmtStream readUInt32];
				CGColorRef underlineColor = [self CGColorWithCCv:ccv];
				[_attributedString setAttribute:kCTUnderlineColorAttributeName toValue:(id)underlineColor startingAtIndex:start endingAtIndex:index];
			}
				
			default:
				[fmtStream skipBytes:[self sizeOfOperandForSprm:sprm]];
				break;
		}
		
	}
	SAFE_RELEASE(fmtStream);
	 
}

-(void)readSectionAttributes
{
	[_tableStream skipToByte:_docInfo.fcPlcfSed];
	uint32_t sedCount = (_docInfo.lcbPlcfSed - 4) / 16;
	uint32_t sedStartLocations[sedCount + 1];
	for(int i = 0; i < (sedCount + 1); i++) {
		sedStartLocations[i] = [_tableStream readUInt32];
	}
	
	for(int i = 0; i < sedCount; i++) {
		[_tableStream skipBytes:2];
		uint32_t fcSepx = [_tableStream readUInt32];
		if(fcSepx != 0xFFFFFFFF) {
			[_wordStream skipToByte:fcSepx];
			uint16_t cbSepx = [_wordStream readUInt16];
			NSData *sepxData = [NSData dataWithBytes:[_wordStream readBytes:cbSepx] length:cbSepx];
			[self applySepx:sepxData fromIndex:sedStartLocations[i] toIndex:sedStartLocations[i + 1]];
		}
		[_tableStream skipBytes:6];
	}
}

-(void)readFontTable
{
	/*Any time a font is referred to in the file, it is referred to by
	 its index in the font table.*/
	
	[_tableStream skipToByte:_docInfo.fcSttbfFfn];
	uint16_t cData = [_tableStream readUInt16];
	[_tableStream skipBytes:2];
	
	for(int i = 0; i < cData; i++) {
		uint8_t cchData = [_tableStream readUInt8];
		[_tableStream skipBytes:39];
		uint32_t fontNameLength = (cchData - 39);
		uint32_t realFontNameLength = fontNameLength;
		uint8_t *xszFfn = [_tableStream readBytes:fontNameLength - 1];
		
		/*Sometimes the xszFfn (font name) is written twice, without warning.
		 The following is a hack-ish way of detecting extra null-terminating
		 characters, thereby eliminating the duplicate name.
		 */
		for(int i = 0; i < (fontNameLength - 3); i++) {
			uint8_t byte = xszFfn[i];
			uint8_t byte2 = xszFfn[i + 1];
			uint8_t byte3 = xszFfn[i + 2];
			if(byte == 0 && byte2 == 0 && byte3 == 0) {
				realFontNameLength = i + 1;
				break;
			}
		}
		if(realFontNameLength != fontNameLength) {
			realFontNameLength++;
		} else {
			realFontNameLength--;
		}
		NSString *fontName = [[NSString alloc] initWithBytes:xszFfn length:realFontNameLength encoding:NSUTF16LittleEndianStringEncoding];
		[_fontNames addObject:fontName];
		SAFE_AUTORELEASE(fontName);
		[_tableStream skipBytes:(cchData - 39) - fontNameLength + 1];
	}
}

-(void)readStylesheet
{
	/*The stylesheet contains the named styles of the document, (eg, "Body", "Header 1", etc.)
	 and describes their formatting properties.
	 Elsewhere in the file, styles are referred to by their index in the stylesheet.
	 
	 */
	_styleGrpprlDatas = [[NSMutableDictionary alloc] init];
	
	[_tableStream skipToByte:_docInfo.fcStshf];
	uint16_t cbStshi = [_tableStream readUInt16];
	uint16_t cstd = [_tableStream readUInt16];
	uint16_t cbSTDBaseInFile = [_tableStream readUInt16];
	[_tableStream skipBytes:8];
	uint16_t ftcAsci = [_tableStream readUInt16];
	NSString *defaultFontName = [_fontNames objectAtIndex:ftcAsci];
	
	CTFontRef defaultFont = CTFontCreateWithName((CFStringRef)defaultFontName, 12.0, NULL);
	[_attributedString setAttribute:kCTFontAttributeName toValue:(id)defaultFont startingAtIndex:0 endingAtIndex:[_attributedString length]];
	CFRelease(defaultFont);
	[_tableStream skipBytes:4];
	
	[_tableStream skipBytes:4];
	//uint16_t cbLSD = [_tableStream readUInt16];
	uint32_t indexInStshi = [_tableStream index] - _docInfo.fcStshf;
	[_tableStream skipBytes:cbStshi - indexInStshi + 2];
	
	for(int i = 0; i < cstd; i++) {
		uint16_t cbStd = [_tableStream readUInt16];
		uint32_t startingIndex = [_tableStream index];
		if(cbStd == 0) {
			continue;
		}
		[_tableStream skipBytes:2];
		uint16_t stkAndIstdBase = [_tableStream readUInt16];
		//uint16_t istdBase = stkAndIstdBase >> 4;
		uint8_t stk = stkAndIstdBase & 0x0F;
		[_tableStream skipBytes:6];
		if(cbSTDBaseInFile == 0x0012) {
			[_tableStream skipBytes:8];
		}
		uint16_t cch = [_tableStream readUInt16];
		[_tableStream skipBytes:((cch + 1) * 2)]; /*We don't need the style name here, but if your application
											 needs it, read these bytes instead of skipping them.*/
		NSMutableData *styleData = [NSMutableData data];
		
		/*This style data is not applied yet, because we don't know where to
		 apply it. For now, just store it in a dictionary as the object whose key
		 is its index in the table. 
		 
		 An NSDictionary is used instead of an NSArray, as there are often empty
		 styles; therefore the indices of the style may not be adjacent.
		 See -readFKPWithData:isPapxFKP: for more info.*/
		switch(stk) {
			//A Paragraph Style
			case 0x01: {
				uint16_t cbUpxPapx = [_tableStream readUInt16];
				[_tableStream skipBytes:2];
				
				uint8_t *upxPapxBytes = [_tableStream readBytes:(cbUpxPapx - 2)];
				[styleData appendBytes:upxPapxBytes length:(cbUpxPapx - 2)];
				if(cbUpxPapx & 1) [_tableStream skipBytes:1];
				
				uint16_t cbUpxChpx = [_tableStream readUInt16];
				uint8_t *upxChpxBytes = [_tableStream readBytes:cbUpxChpx];
				[styleData appendBytes:upxChpxBytes length:cbUpxChpx];
				if(cbUpxChpx & 1) [_tableStream skipBytes:1];
				
				[_tableStream skipToByte:startingIndex + cbStd];
				break;
			}
			
			//A character style
			case 0x02: {
				uint16_t cbUpx = [_tableStream readUInt16];
				uint8_t *upxChpxBytes = [_tableStream readBytes:cbUpx];
				[styleData appendBytes:upxChpxBytes length:cbUpx];
				if(cbUpx & 1) [_tableStream skipBytes:1];
				
				break;
			}
			
			//A table style
			case 0x03: {
				uint16_t cbUpxTapx = [_tableStream readUInt16];
				[_tableStream skipBytes:cbUpxTapx];
				if(cbUpxTapx & 1) [_tableStream skipBytes:1];
				
				uint16_t cbUpxPapx = [_tableStream readUInt16];
				[_tableStream skipBytes:2];
				uint8_t *upxPapxBytes = [_tableStream readBytes:(cbUpxPapx - 2)];
				[styleData appendBytes:upxPapxBytes length:(cbUpxPapx - 2)];
				if(cbUpxPapx & 1) [_tableStream skipBytes:1];
				
				uint16_t cbUpxChpx = [_tableStream readUInt16];
				uint8_t *upxChpxBytes = [_tableStream readBytes:cbUpxChpx];
				[styleData appendBytes:upxChpxBytes length:cbUpxChpx];
				if(cbUpxChpx & 1) [_tableStream skipBytes:1];
				
				break;
			}
			
			//A list style
			case 0x04: {
				uint16_t cbUpx = [_tableStream readUInt16];
				[_tableStream skipBytes:2];
				uint8_t *upxPapxBytes = [_tableStream readBytes:(cbUpx - 2)];
				[styleData appendBytes:upxPapxBytes length:(cbUpx - 2)];
				if(cbUpx & 1) [_tableStream skipBytes:1];
				break;
			}
		}
		[_styleGrpprlDatas setObject:styleData forKey:[NSNumber numberWithInt:i]];
	}
}

-(void)readFormatting
{
	/*Formatting data is stored in Formatted Disk Pages (FKP).
	 The PlcfBteChpx and PlcfBtePapx in the table stream tell
	 us where these Formatted Disk Pages are. 
	 */
	[_tableStream skipToByte:_docInfo.fcPlcfBtePapx];
	uint32_t papxSectorCount = (_docInfo.lcbPlcfBtePapx - 4) / 8;
	[_tableStream skipBytes:(papxSectorCount + 1) * 4];
	
	for(int i = 0; i < papxSectorCount; i++) {
		uint32_t sector = [_tableStream readUInt32];
		NSData *fkpData = [[_wordStream data] subdataWithRange:NSMakeRange((sector * _docInfo.pageSize), _docInfo.pageSize)];
		[self readFKPWithData:fkpData isPapxFKP:YES];
	}
	
	[_tableStream skipToByte:_docInfo.fcPlcfBteChpx];
	uint32_t chpxSectorCount = (_docInfo.lcbPlcfBteChpx - 4) / 8;
	[_tableStream skipBytes:(chpxSectorCount + 1) * 4];
	
	for(int i = 0; i < chpxSectorCount; i++) {
		uint32_t sector = [_tableStream readUInt32];
		NSData *fkpData = [[_wordStream data] subdataWithRange:NSMakeRange((sector * _docInfo.pageSize), _docInfo.pageSize)];
		[self readFKPWithData:fkpData isPapxFKP:NO];
	}
	
}

-(void)readFKPWithData:(NSData *)fkpData isPapxFKP:(BOOL)isPapx
{
	/*An FKP is a 512-byte structure that describes how the text
	 is formatted. The final byte of the FKP, called cb, describes
	 how many different styles this FKP represents.
	 */
	WKDataStream *fkpStream = [[WKDataStream alloc] initWithData:fkpData];
	[fkpStream skipToByte:511];
	uint8_t cb = [fkpStream readUInt8];
	[fkpStream skipToByte:0];
	
	/*The FKP starts with an array of 4-byte integers. Each integer in
	 this array is a file offset where a formatting change occurs in the text.
	 */
	
	uint32_t fcs[cb + 1];
	for(int i = 0; i < (cb + 1); i++) fcs[i] = [fkpStream readUInt32];
	
	/*Following the array of file offsets is an array of 1-byte offsets, describing
	 offsets within this FKP. This array parallels the previous array, such that
	 the formatting change occuring at fcs[i] is described in the FKP at bxs[i] * 2.
	 */
	
	uint8_t bxs[cb];
	for(int i = 0; i < cb; i++) {
		bxs[i] = [fkpStream readUInt8];
		if(isPapx) {
			[fkpStream skipBytes:12];
		}
	}
	
	/*The FKP concludes with the actual formatting data. The data
	 is stored as CHPXs for Character formatting or PAPXs for paragraph
	 formatting. A CHPX or PAPX starts at (bxs[n] * 2), where 0 <= n <= cb.
	 */
	
	for(int i = 0; i < cb; i++) {
		if(isPapx) {
			[fkpStream skipToByte:(bxs[i] * 2)];
			uint8_t cbPapxInFkp = [fkpStream readUInt8];
			if(cbPapxInFkp == 0) {
				cbPapxInFkp = ([fkpStream readUInt8] * 2);
			} else {
				cbPapxInFkp = (cbPapxInFkp * 2) - 1;
			}
			/*PAPXs can include a reference to a style from the stylesheet.
			 This style must be applied before applying the paragraph formatting.
			 */
			uint16_t istd = [fkpStream readUInt16];
			NSData *istdStyleData = [_styleGrpprlDatas objectForKey:[NSNumber numberWithShort:istd]];
			if([istdStyleData length]) {
				[self applyFormatting:istdStyleData fromIndex:[self convertFCtoCP:fcs[i]] toIndex:[self convertFCtoCP:fcs[i + 1]]];
			}
			uint8_t *papxBytes = [fkpStream readBytes:(cbPapxInFkp )];
			NSData *papx = [NSData dataWithBytes:papxBytes length:cbPapxInFkp];
			[self applyFormatting:papx fromIndex:[self convertFCtoCP:fcs[i]] toIndex:[self convertFCtoCP:fcs[i + 1]]];
		} else {
			if(bxs[i] == 0x00) {
				NSData *istdStyleData = [_styleGrpprlDatas objectForKey:[NSNumber numberWithShort:16]];
				if([istdStyleData length]) {
					[self applyFormatting:istdStyleData fromIndex:[self convertFCtoCP:fcs[i]] toIndex:[self convertFCtoCP:fcs[i + 1]]];
				}
				continue;
			}
			[fkpStream skipToByte:(bxs[i] * 2)];
			uint8_t cbChpxInFkp = [fkpStream readUInt8];
			uint8_t *chpxBytes = [fkpStream readBytes:cbChpxInFkp];
			NSData *chpx = [NSData dataWithBytes:chpxBytes length:cbChpxInFkp];
			[self applyFormatting:chpx fromIndex:[self convertFCtoCP:fcs[i]] toIndex:[self convertFCtoCP:fcs[i + 1]]];
		}
	}
	
	SAFE_RELEASE(fkpStream);
}

-(void)readDocumentProperties
{
	[_tableStream skipToByte:_docInfo.fcDop];
	[_tableStream skipBytes:5];
	uint8_t flags8 = [_tableStream readUInt8];
	BOOL fSplHideErrors = (flags8 & 0x01);
	BOOL fGramHideErrors = (flags8 & 0x02);
	[_documentAttributes setObject:[NSNumber numberWithBool:fSplHideErrors] forKey:WKHideSpellingErrorsDocumentAttribute];
	[_documentAttributes setObject:[NSNumber numberWithBool:fGramHideErrors] forKey:WKHideGrammarErrorsDocumentAttribute];
	[_tableStream skipBytes:4];
	uint16_t dxaTab = [_tableStream readUInt16];
	[_documentAttributes setObject:[NSNumber numberWithFloat:((CGFloat)dxaTab / 20.0)] forKey:WKDefaultTabIntervalDocumentAttribute];
	[_tableStream skipBytes:8];
	uint32_t dttmCreated = [_tableStream readUInt32];
	uint32_t dttmModified = [_tableStream readUInt32];
	NSDate *creationDate = [self dateWithDTTM:dttmCreated];
	NSDate *modificationDate = [self dateWithDTTM:dttmModified];
	[_documentAttributes setObject:creationDate forKey:WKCreationTimeDocumentAttribute];
	[_documentAttributes setObject:modificationDate forKey:WKModificationTimeDocumentAttribute];
	[_tableStream skipBytes:54];
	uint16_t flags16 = ([_tableStream readUInt16] & 0x3FFF);
	uint8_t wvkoSaved = (flags16 & 0x07);
	[_documentAttributes setObject:[NSNumber numberWithChar:wvkoSaved] forKey:WKViewModeDocumentAttribute];
	uint16_t pctWwdSaved = (flags16 & 0x0FF8) >> 3;
	[_documentAttributes setObject:[NSNumber numberWithShort:pctWwdSaved] forKey:WKViewZoomDocumentAttribute];
	uint8_t zkSaved = (flags16 & 0xC000) >> 12;
	[_documentAttributes setObject:[NSNumber numberWithChar:zkSaved] forKey:WKAutosizeDocumentAttribute];
	
}

-(void)read
{
	[self readHeaderAndAssembleStreams];
	[self readFileInformationBlock];
	[self readTextContent];
	[self readSectionAttributes];
	[self readFontTable];
	[self readStylesheet];
	[self readFormatting];
	[self readDocumentProperties];
}

#pragma mark -
#pragma mark Utility conversions

//Convert an offset in the file to an index within the actual text
-(uint32_t)convertFCtoCP:(uint32_t)fc
{
	fc -= _docInfo.textStartLocation;
	if(!_docInfo.fcCompressed) {
		fc /= 2;
	}
	return fc;
}

-(uint16_t)sizeOfOperandForSprm:(uint16_t)sprm
{
	uint16_t spra = (uint16_t)floor(sprm / 8192);
	//Pretty much arbitrary...
	switch(spra) {
		case 0x00:
		case 0x01:
			return 1;
			
		case 0x02:
		case 0x04:
		case 0x05:
			return 2;
			
		case 0x03:
			return 4;
			
		case 0x07:
			return 3;
			
		case 0x06:
			return 0;
	}
	return 0;
}

-(CGColorRef)CGColorWithICO:(uint8_t)ico
{
	CGFloat r = 0.0, g = 0.0, b = 0.0;
	switch(ico) {
		case 0x02: b = 1.0; break;
		case 0x03: g = b = 1.0; break;;
		case 0x04: g = 1.0; break;
		case 0x05: r = b = 1.0; break;
		case 0x06: r = 1.0;
		case 0x07: r = g = 1.0; break;
		case 0x08: r = g = b = 1.0; break;
		case 0x09: b = 0.5; break;
		case 0x0A: g = b = 0.5; break;
		case 0x0B: g = 0.5; break;
		case 0x0C: case 0x0D: r = b = 0.5; break;
		case 0x0E: r = g = 0.5; break;
		case 0x0F: r = g = b = 0.5; break;
		case 0x10: r = g = b = 0.75; break;
	}
	return [UIColor colorWithRed:r green:g blue:b alpha:1.0].CGColor;
}

-(CGColorRef)CGColorWithCCv:(uint32_t)ccv
{
	if(ccv & 0xFF000000) {
		return [UIColor blackColor].CGColor;
	}
	uint8_t blue = (ccv & 0x00FF0000) >> 16;
	uint8_t green = (ccv & 0x0000FF00) >> 8;
	uint8_t red = (ccv & 0x000000FF);
	CGFloat blueF = ((CGFloat)blue / 256.0);
	CGFloat greenF = ((CGFloat)green / 256.0);
	CGFloat redF = ((CGFloat)red / 256.0);
	
	return [UIColor colorWithRed:redF green:greenF blue:blueF alpha:1.0].CGColor;
}

-(NSDate *)dateWithDTTM:(uint32_t)dttm
{
	uint8_t mint = (dttm & 0x3F);
	uint8_t hr = (dttm >> 6) & 0x1F;
	uint8_t dom = (dttm >> 11) & 0x1F;
	uint8_t mon = (dttm >> 16) & 0xF;
	uint16_t yr = (dttm >> 20) & 0x01FF;
	
	NSDateComponents *dateComponents = SAFE_AUTORELEASE([[NSDateComponents alloc] init]);
	[dateComponents setMinute:mint];
	[dateComponents setHour:hr];
	[dateComponents setDay:dom];
	[dateComponents setMonth:mon];
	[dateComponents setYear:1900 + yr];
	
	return [[NSCalendar currentCalendar] dateFromComponents:dateComponents];
}

-(CTParagraphStyleRef)createModifiedParagraphStyle:(CTParagraphStyleRef)originalStyle specifier:(CTParagraphStyleSpecifier)specifier newValue:(void *)newValuePtr
{
	static const uint32_t paragraphStyleSpecifierSizes[kCTParagraphStyleSpecifierCount] = {
		sizeof(CTTextAlignment),
		sizeof(CGFloat),
		sizeof(CGFloat),
		sizeof(CGFloat),
		sizeof(CFArrayRef),
		sizeof(CGFloat),
		sizeof(CTLineBreakMode),
		sizeof(CGFloat),
		sizeof(CGFloat),
		sizeof(CGFloat),
		sizeof(CGFloat),
		sizeof(CGFloat),
		sizeof(CGFloat),
		sizeof(CTWritingDirection),
		sizeof(CGFloat),
		sizeof(CGFloat),
		sizeof(CGFloat),
	};
	
	CTParagraphStyleSetting settings[kCTParagraphStyleSpecifierCount];
	for(int i = 0; i < kCTParagraphStyleSpecifierCount; i++) {
		settings[i].spec = i;
		settings[i].valueSize = paragraphStyleSpecifierSizes[i];
		if(specifier == i) {
			settings[i].value = newValuePtr;
			continue;
		}
		void *oldValuePtr;
		CTParagraphStyleGetValueForSpecifier(originalStyle, i, paragraphStyleSpecifierSizes[i], oldValuePtr);
		settings[i].value = oldValuePtr;
	}
	return CTParagraphStyleCreate(settings, kCTParagraphStyleSpecifierCount);
}

#pragma  mark -

-(void)setFailureReason:(NSString *)reason
{
	NSLog(@"%@", reason);
	if(_error) {
		SAFE_RELEASE(_error);
		_error = nil;
	}
	_error = SAFE_RETAIN([NSError errorWithDomain:NSCocoaErrorDomain code:0 userInfo:[NSDictionary dictionaryWithObject:reason forKey:NSLocalizedFailureReasonErrorKey]]);
	
}

#pragma mark -
#pragma mark Public

-(NSAttributedString *)attributedString
{
	return _attributedString;
}

-(NSDictionary *)documentAttributes
{
	return _documentAttributes;
}

-(NSError *)anyErrorDuringReading
{
	return _error;
}
@end
