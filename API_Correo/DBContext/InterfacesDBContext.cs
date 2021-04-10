using API_Correo.DTO;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;

namespace API_Correo.DBContext
{
    public class InterfacesDBContext : DbContext
    {
        public InterfacesDBContext()
        {
        }

        public InterfacesDBContext(DbContextOptions<InterfacesDBContext> options)
           : base(options)
        {
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<EROfisisDTO>().HasNoKey().ToView(null);
        }
    }
}
