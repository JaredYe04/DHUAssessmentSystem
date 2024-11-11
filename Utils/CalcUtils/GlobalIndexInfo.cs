using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
using 考核系统.Mapper;
namespace 考核系统.Utils.CalcUtils
{
    internal class GlobalIndexInfo
    {
        private static List<Entity.Index> indexInfo;
        private static List<Completion> completionColumns;
        private static GlobalIndexInfo instance;
        private GlobalIndexInfo() {
            
            var indexMapper = IndexMapper.GetInstance();
            var sql= "select * from indexes";
            indexInfo = indexMapper.QueryAll(sql);
            var completionMapper = CompletionMapper.GetInstance();
            sql = "select * from completion";
            completionColumns = completionMapper.QueryAll(sql);
        }
        public static GlobalIndexInfo GetInstance()
        {
            if (instance == null)
            {
                instance = new GlobalIndexInfo();
            }
            return instance;
        }

        public int GlobalCompletion(Entity.Index index)
        {
            //返回全校某指标的完成数之和
            return completionColumns.Where(completion => completion.index_id == index.id).Sum(completion => completion.completed);
        }
        public double BasicTheoreticalFullScoreSum//6.单项基础类指标完成度理论满分总分
        {
            get
            {
                return indexInfo.Where(index => index.index_type != "加分类").Sum(index => index.BasicTheoreticalFullScore);
            }
        }
    }
}
