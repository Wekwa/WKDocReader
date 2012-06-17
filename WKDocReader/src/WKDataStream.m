/*
 WKDataStream.m
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
#import "WKDataStream.h"

@implementation WKDataStream

-(id)initWithData:(NSData *)data
{
	self = [super init];
	if(self) {
		_data = [data copy];
		_index = 0;
	}
	return self;
}

-(NSData *)data
{
	return _data;
}

-(UInt8 *)readBytes:(NSUInteger)byteCount
{
	if(buf) free(buf), buf = nil;
	buf = malloc(byteCount);
	[_data getBytes:buf range:NSMakeRange(_index, byteCount)];
	_index += byteCount;
	return buf;
}

-(UInt8)readUInt8
{
	return [self readBytes:1][0];
}

-(UInt16)readUInt16
{
	
	return CFSwapInt16HostToLittle(*(UInt16 *)[self readBytes:2]);
}

-(UInt32)readUInt32
{
	return CFSwapInt32HostToLittle(*(UInt32 *)[self readBytes:4]);
}

-(UInt64)readUInt64
{
	return CFSwapInt64HostToLittle(*(UInt64 *)[self readBytes:8]);
}

-(UInt8 *)peekBytes:(NSUInteger)byteCount
{
	if(buf) free(buf), buf = nil;
	buf = malloc(byteCount);
	[_data getBytes:buf range:NSMakeRange(_index, byteCount)];
	return buf;
	
}

-(UInt8)peekUInt8
{
	return [self peekBytes:1][0];
}

-(UInt16)peekUInt16
{
	return CFSwapInt16HostToLittle(*(UInt16 *)[self peekBytes:2]);
}

-(UInt32)peekUInt32
{
	return CFSwapInt32HostToLittle(*(UInt32 *)[self peekBytes:4]);
}

-(UInt64)peekUInt64
{
	return CFSwapInt64HostToLittle(*(UInt64 *)[self peekBytes:8]);
}

-(void)skipBytes:(NSInteger)byteSkipCount
{
	_index += byteSkipCount;
}

-(void)skipToByte:(NSUInteger)byteIndex
{
	_index = byteIndex;
}

-(NSUInteger)index
{
	return _index;
}

-(BOOL)isAtEnd
{
	return (_index >= [_data length]);
}

@end

