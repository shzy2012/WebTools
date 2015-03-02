public DataTable GetDataSource()
        {
            var report = this.QualityErrorsReportDao.GetAll();
            var user = this.UserManager.GetUnLockedUser();
            var list = from u in user
                       from r in report.Where(x => x.CreatedBy == u.Id.ToString()).DefaultIfEmpty()
                       select new QualityReportObject()
                       {
                           User = u.UserName,
                           Created = r == null ? null : r.Created
                       };

            var transpose = list.ToList();
            var coloum = transpose.GroupBy(x => x.User).Select(c => new { User = c.Key, Created = c.Select(x => x.Created).ToList() }).ToList();
            var table = new DataTable();
            table.TableName = "Transpose";
            foreach (var col in coloum)
            {
                table.Columns.Add(col.User);
            }

            var maxRowCount = coloum.Max(x => x.Created.Count);
            for (int i = 0; i < maxRowCount; i++)
            {
                var row = table.NewRow();
                foreach (var col in coloum)
                {
                    if (col.Created.Count > i)
                    {
                        row[col.User] = col.Created[i];
                    }
                }

                table.Rows.Add(row);
            }

            table.AcceptChanges();
            return table;
        }
