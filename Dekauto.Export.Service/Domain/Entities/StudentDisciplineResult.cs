namespace Dekauto.Export.Service.Domain.Entities
{
    public class StudentDisciplineResult
    {
        public SupplementPlanBucket? PlanBucket { get; set; }

        public string? DisciplineName { get; set; }
        /// <summary>Строка, как приходит из api (число, «зачтено», "х" и тд).</summary>
        public string? Score { get; set; }
        public short? Semester { get; set; }
        public short? Year { get; set; }
        public string? ControlType { get; set; }

        public double? AudHours { get; set; }
        public double? CreditUnits { get; set; }

        public int? PlanOrder { get; set; }

        public bool RequiresManualValidation { get; set; }

        /// <summary>Строка только из карточки, без пары в плане — отдельный блок в конце.</summary>
        public bool IsCardOnlyUnmatchedPlan { get; set; }
    }
}
