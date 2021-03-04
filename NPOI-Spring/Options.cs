namespace NPOI_Spring
{
    public static class Options
    {
        public enum TypeOfDoorOperation
        {
            MechanicalIndustrial1Inch = 1,
            
            ElectricIndustrial1Inch = 2,
            
            MechanicalIndustrial1And1_4Inches = 3,
            
            ElectricIndustrial1And1_4Inches = 4,
            
            MechanicalResidential = 5,
            
            ElectricResidential = 6,
            
            NotWithACurrentDrum = 7
        }

        public enum Calculation
        {
            AlcomexStockSprings = 1,
            
            OnLength = 2
        }

        public enum Cycles
        {
            TenThousand = 1,
            
            FifteenThousand = 2,
            
            TwentyThousand = 3,
            
            TwentyFiveThousand = 4,
            
            ThirtyFiveThousand = 5,
            
            FiftyThousand = 6,
            
            OneHundredThousand = 7,
            
            HeavyDuty = 8
        }

        public enum CableDrum
        {
            TwentyThreeSixteen = 1,
            
            ThirtySixteen = 2,
            
            TwentyThreeTwenty = 3,
            
            ThirtyTwenty = 4
        }

        public enum SpringDInside
        {
            FiftyOne = 1,
            
            SixtySix = 2,
            
            NinetyFive = 3,
            
            OneFiftyTwo = 4,
            
            Program = 5
        }

        public enum SprintNumber
        {
            One = 1,
            
            Two = 2,
            
            Three = 3,
            
            Four = 4
        }
    }
}