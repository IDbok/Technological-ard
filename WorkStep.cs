
namespace Technological_card
{
    public class WorkStep
    {
        static Dictionary<int,float> stagesDic = new Dictionary<int,float>(); // conteins number of stage and time of execution
        int num { get; set; }
        string description { get; set; }
        string? staff { get; set; }
        float stepExecutionTime { get; set; }
        int stage { get; set; }
        float stageExecutionTime { get; set; }
        float machineExecutionTime { get; set; }
        string protections { get; set; }
        string? comments { get; set; }

        public WorkStep()
        {
        }
        public WorkStep(int num,string description, string? staff, float stepExecutionTime, int stage, float stageExecutionTime,
            float machineExecutionTime, string protections, string? comments)
        {
            this.num = num;
            this.description = description;
            this.staff = staff;
            this.stepExecutionTime = stepExecutionTime;
            this.stage = stage;
            // TODO: add stage to stagesDic
            this.stageExecutionTime = stageExecutionTime;
            this.machineExecutionTime = machineExecutionTime;
            this.protections = protections;
            this.comments = comments;
        }

        public int Num { get { return num; } set { num = value; } }
        public string Description { get { return description; } set { description = value; } }
        public string? Staff { get { return staff; } set { staff = value; } }
        public float StepExecutionTime { get { return stepExecutionTime; } set { stepExecutionTime = value; } }
        public int Stage { get { return stage; } set { stage = value; } }
        public float StageExecutionTime { get { return stageExecutionTime; } set { stageExecutionTime = value; } }
        public float MachineExecutionTime { get { return machineExecutionTime; } set { machineExecutionTime = value; } }
        public string Protections { get { return protections; } set { protections = value; } }
        public string? Comments { get { return comments; } set { comments = value; } }

        public static int GetLastStageNum() => !stagesDic.Any() ? 0: stagesDic.Last().Key;
        public static float GetLastStageTime() => stagesDic.Last().Value;
        public static void AddStage(int stageNum, float stageTime) => stagesDic.Add(stageNum, stageTime);
       

    }

}
