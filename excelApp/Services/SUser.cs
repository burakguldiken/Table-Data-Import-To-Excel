using kanbanboard.core.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace excelApp
{
    public class SUser : IUser
    {
        IDbContext dbContext;
        public SUser(IDbContext dbContext)
        {
            this.dbContext = dbContext;
        }

        public List<User> GetUsers()
        {
            string sql = @"SELECT * FROM user";
            return dbContext.ExecuteCommand<User>(sql).ToList();
        }
    }
}
