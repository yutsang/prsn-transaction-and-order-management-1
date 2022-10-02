# prsn-transaction-and-order-management-1
VBA Tasks - 20220302  

## Introduction
The Userform OrderForm in the file is designed to calculate each order value at CityU Shop and maintain a cumulative order value report. You must not hard coding any product names (such as “French Fries”, etc.), any product prices (such as “$25.0” or 25.0, etc.), or any ‘Items in Stock’ values (such as 30) in your submitted code. These types of information must be determined at run time with appropriate code. You are also not allowed to use any information contained in the ‘Products Sheet’ worksheet. Do not change the contents in the ‘Products Sheet’. You are asked to provide code for the following tasks.   

### Task 1
Before OrderForm is drawn to the screen at run time, disable the Frame control fmeQuantity programmatically, but keep fmeQuantity visible.  

### Task 2
An order may contain one or more products. The user is allowed to select only one product at a time from the ListBox control lstProducts for an order. When a product in lstProducts is selected:  
  * If the value of ‘Items in Stock’ of the selected product is larger than 0, then:
    * Set the currently displayed value of ‘Items in Stock’ of the selected product as the maximum value of the ScrollBar control sbrQuantity.
    *	Set the current value of sbrQuantity to 1
    *	Enable fmeQuantity
  *	If the value of ‘Items in Stock’ of the selected product is 0, then:
    *	Display this message in a system dialog box: ‘Selected product is out of stock. Please select another product.’
    *	Disable fmeQuantity, but keep it visible.

### Task 3  
The user will use sbrQuantity to specify the order quantity of the selected product at run time. Whenever the sbrQuantity value is changed, display its latest value programmatically in the caption of the Label control lblQuantity as ‘X items’ where X is the current sbrQuantity value.  

### Task 4 
After a product has been selected and the order quantity (default order quantity is 1) has been specified at run time, the user may click the CommandButton control btnOrder to process the selection. The user will repeat this process for each selected product of the same order. At each time when the user clicks btnOrder for a selected product:  
  *	Add the name of the selected product to the bottom of the first column of the ListBox control lstOrdered. The same product may appear more than once in lstOrdered if selected more than once in the same order. There is no need to group the identical products in lstOrdered. 
  *	Add the sbrQuantity value of the selected product to the bottom of the second column of lstOrdered.
  *	Add the order value of the selected product to the bottom of the third column of lstOrdered. Do not include the $ symbol in the column. The order value of a chosen product equals its price times the current sbrQuantity value.
  *	Increment the caption of Label control lblTotal by the order value of the selected product as computed in Task 4c. The caption must be displayed with a $ symbol at the beginning and always in exactly one decimal place (such as $125.0 or $125.8).
  *	Reduce the displayed value of ‘Items in Stock’ of the selected product in lstProducts by the current sbrQuantity value.
  *	Set the value of sbrQuantity to 1.
  *	Disable fmeQuantity.
  *	Deselect the selected product in lstProducts.  

{Hint: For example, OrderForm may look like the following at runtime}  
* Before any selection:  
  <img width="271" alt="image" src="https://user-images.githubusercontent.com/19373417/193460538-c4b9e01d-2414-41f8-867a-a37fda759658.png">
* After the completion of three selections, which includes 10 Hamburger, 1 Chicken Nugget and 4 Chicken Nuggets:  
  <img width="272" alt="image" src="https://user-images.githubusercontent.com/19373417/193460557-53459be5-7ea9-4e23-aa98-e647ddf245cc.png">  
  
### Task 5  
The user may select one or more products in lstOrdered and then remove them from lstOrdered. When btnRemove is clicked, repeat the following actions for each selected product in lstOrdered:  
  * Increment the respective value of ‘Items in Stock’ of the selected product in lstProducts by the current quantity of the selected product in lstOrdered. 
  * Decrease the displayed total value as shown in the caption of lblTotal by the order value of the selected product in lstOrdered. The updated caption must still be displayed with a $ symbol at the beginning and always in exactly one decimal place.
  * Remove the selected product from lstOrdered. The updated list in lstOrdered must not contain blank lines between any two listed products.  
{Hint: For example, OrderForm may look like the following at runtime}  
* Before any removal:  
  <img width="269" alt="image" src="https://user-images.githubusercontent.com/19373417/193460611-012df4a6-e063-409e-a5f3-92e5bef35bfe.png">  
* After 10 Hamburger and 4 Chicken Nuggets are removed:  
  <img width="270" alt="image" src="https://user-images.githubusercontent.com/19373417/193460623-b55404d0-0e97-4229-9a40-0e3ff6e2702b.png">  
  
### Task 6  
When the user clicks the CommandButton control btnProcOrder:  
*	Accumulate all the order values generated so far in each runtime session.
*	Delete all items in lstOrdered.
*	Reset the caption of lblTotal to ‘$0.0’.

### Task 7  
When the user checks the CheckBox control chkAccValue:  
*	Show a message dialog box with the title ‘CityU Shop’ and this message in the dialog box: “The accumulated order value is: $x,xxx.x”, where x,xxx.x is the accumulated order value of all orders in the same session with thousand commas as needed and in exactly one decimal place (such as $1,234.5, or $123.0). The dialog box should look like this:  
  <img width="158" alt="image" src="https://user-images.githubusercontent.com/19373417/193460666-5c92ec0f-423f-40ba-88e0-3af0dbefd1e0.png">  
*	Uncheck chkAccValue programmatically.  

### Task 8
When the user clicks the CommandButton control btnQuit:  
*	Show this message in a dialog box with ‘Yes’ and ‘No’ buttons: “Do you want to quit?”.
* If the user clicks the ‘Yes’ button on the dialog box, unload the userform.

### Task 9
Provide code so that the UserForm OrderForm will be shown automatically when the workbook containing OrderForm is opened.  



	 
