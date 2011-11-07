using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SWManMonth
{
    class Model
    {
        private String modelCode;
        private List<CA> modelCAs;
        private CA firstCA;
        private Int32 numberModelCAs;

        public Model()
        {
            this.numberModelCAs = 0;
            modelCAs = new List<CA>();
        }

        public String ModelCode
        {
            get { return this.modelCode; }
            set { this.modelCode = value.ToString(); }
        }

        public List<CA> ModelCas
        {
            get { return this.modelCAs; }
        }

        public CA ModelCAs
        {
            get
            {
                firstCA = new CA();
                foreach(CA ca in modelCAs){
                    firstCA.CarrierName = ca.CarrierName;
                    firstCA.Country = ca.Country;
                    firstCA.MediumManMonth = ca.MediumManMonth;
                }
                return firstCA;
            }
            set 
            {
                if (value.GetType() == typeof(CA))
                {
                    CA ca = new CA();

                    ca.CarrierName = value.CarrierName;
                    ca.Country = value.Country;
                    ca.MediumManMonth = value.MediumManMonth;

                    modelCAs.Add(ca);
                }
                
            }
        }

    }
}
