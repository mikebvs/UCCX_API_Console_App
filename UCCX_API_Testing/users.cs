using System;
using System.Collections.Generic;
using System.Text;

namespace UCCX_API_Testing
{
    class User
    {
        public User(string id, string firstname, string lastname, string extension, string refgroupname, string refgroupurl)
        {
            Id = id;
            FirstName = firstname;
            LastName = lastname;
            Extension = extension;
            refGroupName = refgroupname;
            refGroupURL = refgroupurl;
        }
        string Id { get; set; }
        string FirstName { get; set; }
        string LastName { get; set; }
        string Extension { get; set; }
        string refGroupName { get; set; }
        string refGroupURL { get; set; }
        public void Info()
        {
            Console.WriteLine("Id: {0}\nFirst Name: {1}\nLast Name: {2}\nExtension: {3}\nGroup Name: {4}\nGroup URL: {5}", Id, FirstName, LastName, Extension, refGroupName, refGroupURL);
        }
    }
    //public User() { }
}
