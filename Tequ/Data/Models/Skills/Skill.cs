namespace Tequ.Data.Models.Skills
{
    public class Skill
    {
        public string SkillType { private set; get; }
        public int Level { private set; get; }
        public long Experience { private set; get; }

        public Skill() {
            this.SkillType = "null";
            this.Level = 0;
            this.Experience = 0;
        }

        public Skill(string skillType) {
            this.SkillType = skillType;
            this.Level = 0;
            this.Experience = 0;
        }

        public Skill(string skillType, long amount) {
            this.SkillType = skillType;
            AddExperience(amount);
        }

        public void AddExperience(long amount) {
            this.Experience += amount;

            while (this.Experience >= GetNextLevelExperience()) {
                this.Experience -= GetNextLevelExperience();
                this.Level += 1;
            }
        }

        public long GetNextLevelExperience() {
            return 0;
        }
    }
}
