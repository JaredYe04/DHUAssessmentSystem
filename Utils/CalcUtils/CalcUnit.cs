using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
namespace 考核系统.Utils.CalcUtils
{
    internal class CalcUnit
    {
        private Index index;
        private Department department;
        private DeptAnnualInfo deptAnnualInfo;
        private Completion completion;
        public CalcUnit(Index index, Department department,DeptAnnualInfo deptAnnualInfo,Completion completion)
        {
            this.index = index;
            this.department = department;
            this.deptAnnualInfo = deptAnnualInfo;
            this.completion = completion;
        }   

        public double BasicCompletionScore//1.基础类完成度得分
        {
            get
            {
                if (index.index_type != "加分类") return 0;
                var completionRate = completion.completion_rate;
                if (completionRate < 1.0) { }
                else completionRate = 1.0;

                return 100 * index.weight1 * index.weight2 * completionRate;
            }
        }
        public double BonusCompletionScore//2.加分类类完成度得分
        {
            get
            {
                if (index.index_type != "加分类") return 0;
                var completionRate = completion.completion_rate;
                if (completionRate < 0.6) completionRate = 0;
                else if (completionRate < 1.0) { }
                else if (completionRate < 2.0) completionRate = 1.0;
                else completionRate *= 0.6;
                return 100 * index.weight1 * index.weight2 * completionRate;
            }
        }
        public double BasicCompletionScorePerCapita//3.基础类人均贡献度得分
        {
            get
            {
                if (index.index_type == "加分类") return 0;
                var completed = completion.completed;
                var population = (double)(deptAnnualInfo.dept_population);

                var globalDeptInfo = GlobalDeptInfo.GetInstance();
                var globalPopulation = (double)(globalDeptInfo.GlobalPopulation(deptAnnualInfo.year));

                var globalIndexInfo = GlobalIndexInfo.GetInstance();
                var globalCompletion = (double)(globalIndexInfo.GlobalCompletion(index));
                if (globalCompletion==0 || population==0) return 0;//todo，按理来说这里应该抛出异常
                return 100 * index.weight1 * index.weight2 *
                    (completed / population) * (globalPopulation / globalCompletion);
            }
        }
        public double SensitivityScore//4.敏感性指标得分
        {
            get
            {
                if(index.sensitivity == 0) return 0;//不以enable_sensitivity为判断标准，只要sensitivity不为0则算
                var completed = completion.completed;
                var population = (double)(deptAnnualInfo.dept_population);

                var globalDeptInfo = GlobalDeptInfo.GetInstance();
                var globalPopulation = (double)(globalDeptInfo.GlobalPopulation(deptAnnualInfo.year));

                var globalIndexInfo = GlobalIndexInfo.GetInstance();
                var globalCompletion = (double)(globalIndexInfo.GlobalCompletion(index));
                if (globalCompletion == 0 || population == 0) return 0;//todo，按理来说这里应该抛出异常
                return 100 * index.sensitivity * 
                    (completed / population) / (globalCompletion / globalPopulation);
            }
        }
        public double BasicTheoreticalFullScore//5.单项基础类指标完成度理论满分
        {
            get
            {
                return index.BasicTheoreticalFullScore;
            }
        }
        //6.见GlobalIndexInfo

        public double ObjectiveScoreSum//7.客观分合计,一个学院对一个指标的客观分合计
        {
            get
            {
                return BasicCompletionScore + BonusCompletionScore + BasicCompletionScorePerCapita + SensitivityScore;
            }
        }
        public double WeightedObjectiveScoreSum//8.加权客观分合计
        {
            get
            {
                return (BasicCompletionScore + BonusCompletionScore) * 0.25 +
                    (BasicCompletionScorePerCapita + SensitivityScore) * 0.25;
            }
        }
    }
}
