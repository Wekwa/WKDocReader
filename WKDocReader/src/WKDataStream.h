/*
 WKDataStream.h
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

/*An easy data-reading class that doesn't make you deal with pointers like NSData or NSStream does*/

@interface WKDataStream : NSObject {
    NSData *_data;
	NSUInteger _index;
	UInt8 *buf;
}

-(id)initWithData:(NSData *)data;
-(NSData *)data;

-(UInt8 *)readBytes:(NSUInteger)byteCount;
-(UInt8)readUInt8;
-(UInt16)readUInt16;
-(UInt32)readUInt32;
-(UInt64)readUInt64;

/*Peeking does not increase the index, unlike reading.*/

-(UInt8 *)peekBytes:(NSUInteger)byteCount;
-(UInt8)peekUInt8;
-(UInt16)peekUInt16;
-(UInt32)peekUInt32;
-(UInt64)peekUInt64;

-(void)skipBytes:(NSInteger)byteSkipCount;
-(void)skipToByte:(NSUInteger)byteIndex;

-(BOOL)isAtEnd;

-(NSUInteger)index;

@end
