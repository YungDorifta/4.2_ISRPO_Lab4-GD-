//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Лаб4AJAX
{
    using System;
    using System.Collections.Generic;
    
    public partial class BooksReaders
    {
        public int ID { get; set; }
        public int ID_book { get; set; }
        public int ID_reader { get; set; }
    
        public virtual Books Books { get; set; }
        public virtual Readers Readers { get; set; }
    }
}
