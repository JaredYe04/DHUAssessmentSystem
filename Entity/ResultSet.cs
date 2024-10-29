using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Table;

namespace 考核系统.Entity
{
    class ResultSet 
    {
        double basicCompletionScore = double.NaN;//基础类完成度得分
        double bonusCompletionScore = double.NaN;//加分类完成度得分
        double basicPerCapitaScore = double.NaN;//基础类人均贡献度得分
        double sensitivityIndexScore = double.NaN;//敏感性指标得分


        double basicTheoreticalFullScore = double.NaN;//基础类指标完成度理论满分

        public static ResultSet Calculate(string departmentId,string indexId)
        {
            var department = DepartmentTable.GetDepartmentById(departmentId);
            var index = IndexTable.GetIndexById(indexId);

            var resultSet = new ResultSet(0, 0, 0, 0, 0);
            //todo,根据具体的计算公式计算结果

            return resultSet;
        }





        //get set constructor
        public double BasicCompletionScore
        {
            get { return basicCompletionScore; }
            set { basicCompletionScore = value; }
        }
        public double BonusCompletionScore
        {
            get { return bonusCompletionScore; }
            set { bonusCompletionScore = value; }
        }
        public double BasicPerCapitaScore
        {
            get { return basicPerCapitaScore; }
            set { basicPerCapitaScore = value; }
        }
        public double SensitivityIndexScore
        {
            get { return sensitivityIndexScore; }
            set { sensitivityIndexScore = value; }
        }

        public double ObjectiveTotalScore//客观分合计
        {
            get { return basicCompletionScore + bonusCompletionScore + basicPerCapitaScore + sensitivityIndexScore; }
        }
        public double ObjectiveWeightedTotalScore//客观分加权合计
        {
            get { return (basicCompletionScore + bonusCompletionScore)*0.25 
                    + (basicPerCapitaScore + sensitivityIndexScore)*0.75;
            }//todo,这里的权重需要根据具体情况调整
        }


        public double BasicTheoreticalFullScore
        {
            get { return basicTheoreticalFullScore; }
            set { basicTheoreticalFullScore = value; }
        }

        private ResultSet(double basicCompletionScore, double bonusCompletionScore, double basicPerCapitaScore, double sensitivityIndexScore, double basicTheoreticalFullScore)
        {
            this.basicCompletionScore = basicCompletionScore;
            this.bonusCompletionScore = bonusCompletionScore;
            this.basicPerCapitaScore = basicPerCapitaScore;
            this.sensitivityIndexScore = sensitivityIndexScore;
            this.basicTheoreticalFullScore = basicTheoreticalFullScore;
        }


    }
}
