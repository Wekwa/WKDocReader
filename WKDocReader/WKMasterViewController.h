#import <UIKit/UIKit.h>

@class WKDetailViewController;

@interface WKMasterViewController : UITableViewController {
	NSMutableArray *documents;
}

@property (strong, nonatomic) WKDetailViewController *detailViewController;

@end
