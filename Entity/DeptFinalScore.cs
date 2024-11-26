using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    internal class DeptFinalScore
    {
        public Department department { get; set; }
        public DeptAnnualInfo deptAnnualInfo { get; set; }
        public double BasicCompletionScoreSum { get; set; }
        public double BonusCompletionScoreSum { get; set; }
        public double BasicCompletionScorePerCapitaSum { get; set; }
        public double SensitivityScoreSum { get; set; }
        public double BasicTheoreticalFullScoreSum {  get; set; }

        public DeptFinalScore(Department department,DeptAnnualInfo deptAnnualInfo,double basicCompletionScoreSum, double bonusCompletionScoreSum, double basicCompletionScorePerCapitaSum, double sensitivityScoreSum, double basicTheoreticalFullScoreSum)
        {
            this.department = department;
            this.deptAnnualInfo = deptAnnualInfo;
            BasicCompletionScoreSum = basicCompletionScoreSum;
            BonusCompletionScoreSum = bonusCompletionScoreSum;
            BasicCompletionScorePerCapitaSum = basicCompletionScorePerCapitaSum;
            SensitivityScoreSum = sensitivityScoreSum;
            BasicTheoreticalFullScoreSum = basicTheoreticalFullScoreSum;
        }

        public double ObjectiveScoreSum
        {
            get
            {
                return BasicCompletionScoreSum + BonusCompletionScoreSum + BasicCompletionScorePerCapitaSum + SensitivityScoreSum;
            }
        }

        public double WeightedCompletionScoreSum
        {
            get
            {
                return 0.25 * (BasicCompletionScoreSum + BonusCompletionScoreSum);
            }
        }

        public double WeightedContributiveScoreSum
        {
            get
            {
                return 0.25 * (BasicCompletionScorePerCapitaSum + SensitivityScoreSum);
            }
        }
        public double WeightedObjectiveScoreSum
        {
            get
            {
                return WeightedCompletionScoreSum + WeightedContributiveScoreSum;
            }
        }
        public double NormalizedObjectiveScoreSum(double highestObjectiveScoreSum)
        {
            return WeightedObjectiveScoreSum / highestObjectiveScoreSum * 50;//最高分为50，其他分数按比例缩放
        }
        public double PunishRate
        {
            get
            {
                return deptAnnualInfo.dept_punishment;
            }
        }
        public double Punishment
        {
            get
            {
                return BasicTheoreticalFullScoreSum * PunishRate;
            }
        }
        public double FinalScore
        {
            get
            {
                return BasicCompletionScoreSum + BasicCompletionScorePerCapitaSum - Punishment;
            }
        }
    }
}
