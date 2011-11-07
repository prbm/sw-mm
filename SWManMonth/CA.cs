using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SWManMonth
{
    class CA
    {
        private String carrierName;
        private String country;
        private Double numberWorkingPeople;
        private Double totalManMonth;
        private Double mediumManMonth;
        private Double totalWorkingHours;
        private Double mediumWorkingHours;

        public String CarrierName
        {
            get
            {
                return this.carrierName;
            }
            set
            {
                this.carrierName = value.ToString();
            }
        }

        public String Country
        {
            get
            {
                return this.country;
            }
            set
            {
                this.country = value.ToString();
            }
        }

        public Double NumberWorkingPeople
        {
            get
            {
                return this.numberWorkingPeople;
            }
            set
            {
                this.numberWorkingPeople = Double.Parse(value.ToString());
            }
        }

        public Double TotalManMonth
        {
            get
            {
                return this.totalManMonth;
            }
            set
            {
                this.totalManMonth = Double.Parse(value.ToString());
            }
        }

        public Double MediumManMonth
        {
            get
            {
                return this.mediumManMonth;
            }
            set
            {
                this.mediumManMonth = Double.Parse(value.ToString());
            }
        }

        public Double TotalWorkingHours
        {
            get
            {
                return this.totalWorkingHours;
            }
            set
            {
                this.totalWorkingHours = Double.Parse(value.ToString());
            }
        }

        public Double MediumWorkingHours
        {
            get
            {
                return this.mediumWorkingHours;
            }
            set
            {
                this.mediumWorkingHours = Double.Parse(value.ToString());
            }
        } 
    }
}
