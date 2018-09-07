# CADOFields

Contains all the **Field** objects of a **Recordset** or **Record** object.

A **Recordset** object has a **Fields** collection made up of **Field** objects. Each **Field** object corresponds to a column in the **Recordset**. You can populate the **Fields** collection before opening the **Recordset** by calling the **Refresh** method on the collection.

See the **Field** object topic for a more detailed explanation of how to use **Field** objects.

The **Fields** collection has an **Append** method, which provisionally creates and adds a Field object to the collection, and an Update method, which finalizes any additions or deletions.

A **Record** object has two special fields that can be indexed with **FieldEnum** constants. One constant accesses a field containing the default stream for the **Record**, and the other accesses a field containing the absolute URL string for the **Record**.

Certain providers (for example, the Microsoft OLE DB Provider for Internet Publishing) may populate the **Fields** collection with a subset of available fields for the **Record** or **Recordset**. Other fields will not be added to the collection until they are first referenced by name or indexed by your code.

If you attempt to reference a nonexistent field by name, a new **Field** object will be appended to the **Fields** collection with a **Status** of **adFieldPendingInsert**. When you call Update, ADO will create a new field in your data source if allowed by your provider.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Append](#Append) | Appends an object to a collection. |
| [Attach](#Attach) | Attaches a reference to a **Field** object to the class. |
| [CancelUpdate](#CancelUpdate) | Cancels any changes made to the **Fields** collection of a **Record** object before calling the **Update** method. |
| [Count](#Count) | Retrieves the number of objects in the **Fields** collection. |
| [Delete_](#Delete_) | Deletes an object from the **Fields** collection. |
| [Item](#Item) | Indicates a specific member of the **Fields** collection, by name or ordinal number. |
| [Resync](#Resync) | Resynchronizes the contents of the **Fields** collection. |
| [Refresh](#Refresh) | Refreshes the contents of the **Fields** collection. |
| [Update](#Update) | Saves any changes you make to the current **Fields** collection of a **Record** object. |
