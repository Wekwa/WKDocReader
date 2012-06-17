/*
 WKCoreTextView.m
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


#import "WKCoreTextView.h"

#define PADDING 3.0f

@implementation WKCoreTextView

@synthesize attributedString;

-(void)setAttributedString:(NSAttributedString *)newAttributedString
{
	[attributedString release];
	attributedString = [newAttributedString copy];
	[self setNeedsDisplay];
}

-(void)drawRect:(CGRect)rect
{
	CGContextRef ctx = UIGraphicsGetCurrentContext();
	
	[[UIColor whiteColor] set];
	CGContextFillRect(ctx, rect);
	
	
	CGContextTranslateCTM(ctx, 0, self.bounds.size.height);
	CGContextScaleCTM(ctx, 1.0, -1.0);
	CGContextSetTextMatrix(ctx, CGAffineTransformMakeScale(1.0, 1.0));
	
	UIBezierPath *path = [UIBezierPath bezierPathWithRect:CGRectMake(PADDING, (-PADDING - 1), self.bounds.size.width - (PADDING * 2), self.bounds.size.height - (PADDING))];
	CTFramesetterRef framesetter = CTFramesetterCreateWithAttributedString((CFAttributedStringRef)attributedString);
	frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path.CGPath, NULL);
	CTFrameDraw(frame, ctx);
	CFRelease(framesetter);
	CFRelease(frame);
}



@end
