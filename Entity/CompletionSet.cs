using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;

namespace 考核系统.Entity
{
    class CompletionSet
    {
        Dictionary<int, int> targets;//key是年份，value是目标完成数
        int completion;//今年的完成数
        double completionRate;//今年的完成率

        Dictionary<int, int> Targets
        {
            get { return targets; }
            set { targets = value; }
        }

        int Completion
        {
            get { return completion; }
            set { completion = value; }
        }

        double CompletionRate
        {
            get { return completionRate; }
            set { completionRate = value; }
        }

        public CompletionSet(Dictionary<int, int> targets, int completion,double completionRate)
        {
            this.targets = targets;
            this.completion = completion;
            this.completionRate = completionRate;
        }
    }
}
