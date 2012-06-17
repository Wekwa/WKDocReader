//
//  WKMasterViewController.h
//  WKDocReader
//
//  Created by Wyatt Kaufman on 6/17/12.
//  Copyright (c) 2012 __MyCompanyName__. All rights reserved.
//

#import <UIKit/UIKit.h>

@class WKDetailViewController;

@interface WKMasterViewController : UITableViewController

@property (strong, nonatomic) WKDetailViewController *detailViewController;

@end
