using System.Collections.Generic;
using Tequ.Data.Models.Items;
using Tequ.Data.Models.Skills;

namespace Tequ.Data.Models.Players
{
    public class Player
    {
        public string Name;
        private Dictionary<string, CompositeItem> _equipment = new Dictionary<string, CompositeItem>();
        private Dictionary<string, Skill> _skills = new Dictionary<string, Skill>();

        public CompositeItem GetEquipment(string equipmentType) {
            if (this._equipment.ContainsKey(equipmentType)) {
                return this._equipment[equipmentType];
            } else {
                return this._equipment[equipmentType];
            }
        }

        public void AddEquipment(string equipmentType, CompositeItem items) {
            if (!this._equipment.ContainsKey(equipmentType)) {
                this._equipment.Add(equipmentType, items);
            } else {
                throw new System.Exception("Equipment type: " + equipmentType + " already exists!");
            }
        }

        public Skill GetSkill(string skillType) {
            if (this._skills.ContainsKey(skillType)) {
                return this._skills[skillType];
            } else {
                this._skills.Add(skillType, new Skill());
                return this._skills[skillType];
            }
        }

        public void AddSkill(string skillType, Skill skill) {
            if (!this._skills.ContainsKey(skillType)) {
                this._skills.Add(skillType, skill);
            } else {
                throw new System.Exception("Skill type: " + skillType + " already exists!");
            }
        }
    }
}
