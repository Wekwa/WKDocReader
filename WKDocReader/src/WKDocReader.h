/*
 WKDocReader.h
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


#import <Foundation/Foundation.h>

//Some additional attributes that can be used as attribute keys with NSAttributedString. (If you are drawing the NSAttributedString with CoreText, you will have to draw these manually.)
extern NSString *const WKColumnCountAttributeName; //An NSNumber indicating the number of columns.
extern NSString *const WKPageOrientationAttributeName; //An NSNumber: 0 = portrait, 1 = landscape.
extern NSString *const WKPageWidthAttributeName; //An NSNumber indicating the page width in points.
extern NSString *const WKPageHeightAttributeName; //An NSNumber indicating the page height in points.
extern NSString *const WKLeftMarginAttributeName; //An NSNumber indicating the left margin width in points.
extern NSString *const WKRightMarginAttributeName; //An NSNumber indicating the right margin width in points.
extern NSString *const WKTopMarginAttributeName; //An NSNumber indicating the top margin height in points.
extern NSString *const WKBottomMarginAttributeName; //An NSNumber indicating the bottom margin height in points.
extern NSString *const WKBackgroundColorAttributeName; //A CGColorRef indicating the background color of the text.

//Valid keys for the documentAttributes dictionary
extern NSString *const WKReadOnlyDocumentAttribute; //An NSNumber, as a BOOL, indicating if the document is read-only.
extern NSString *const WKHideSpellingErrorsDocumentAttribute; //An NSNumber, as a BOOL, indicating if spelling errors should be hidden.
extern NSString *const WKHideGrammarErrorsDocumentAttribute; //An NSNumber, as a BOOL, indicating if grammar errors should be hidden.
extern NSString *const WKDefaultTabIntervalDocumentAttribute; //An NSNumber indicating the default tab interval in points.
extern NSString *const WKCreationTimeDocumentAttribute; //An NSDate indicating when the document was created.
extern NSString *const WKModificationTimeDocumentAttribute; //An NSDate indicating when the document was last modified.
extern NSString *const WKViewModeDocumentAttribute; //An NSNumber indicating the last view mode: 0 = normal, 1 = page layout
extern NSString *const WKViewZoomDocumentAttribute; //An NSNumber between 10 and 500 indicating the zoom percent of the document view.
extern NSString *const WKAutosizeDocumentAttribute; //An NSNumber indicating if the document was displayed with auto-zooming. (0 = no, 1 = fit full page, 2 = fit page width)

@interface WKDocReader : NSObject

-(id)initWithDocFormatData:(NSData *)data;
-(id)initWithContentsOfFile:(NSString *)filePath;

-(NSAttributedString *)attributedString;
-(NSDictionary *)documentAttributes;
-(NSError *)anyErrorDuringReading;


@end
