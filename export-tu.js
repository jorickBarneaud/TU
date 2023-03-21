const XLSX = require('xlsx');
const path = require('path');
const TuTemplate = path.join(__dirname, '../assets/Template_TU_IBM_Interactive.xlsx');
const { Tu, Collaborator } = require("../database/models");
const fs = require('fs');


/**
 * Fonction qui charge le modèle Excel existant
 * @returns {Object} le workbook de l'Excel
 */
 async function loadWorkbook() {
  try {
    const workbook = XLSX.readFile(TuTemplate);
    return workbook;
  } catch (error) {
    throw new Error ('Impossible de charger le template : ' + error);
  }
}


/**
 * Fonction qui récupère une feuille de calcul à partir d'un workbook
 * @param {Object} workbook - le workbook de l'Excel
 * @param {string} sheetName - le nom de la feuille de calcul
 * @returns {Object} la feuille de calcul
 */
 function getWorksheet(workbook, sheetName) {
  try {
    const worksheet = workbook.Sheets[sheetName];
    return worksheet;
  } catch (error) {
    throw new Error (`Impossible de récupérer la feuille de calcul "${sheetName}" : ` + error);
  }
}


async function exportTuFile(res) {
  try {
    const tuData = await Tu.findAll();
    if(!tuData){
      throw new Error ('data is null')
    }
    const CollaboratorsInfos = await Collaborator.findAll();
    if(!CollaboratorsInfos){
      throw new Error ('collaborators is null')
    }

    //Hybrides Cloud Service
    const HybridesCloudService = []
    const HybridesCloudManagement = []
    const HybridesCloudTransformation = []
    //BTS#
    const BTS = []
    const CustomerTransformation = []
    const EntrepriseStartegy = []
    const SalesForce = []
    const FinanceChainTransformation = []
    const DataTechnologyTransformation = []
    const IndusrtyTransformation = []
    const TalentTransformation = []
    let ExcelSynthLine = []

    // trie des collaborateurs en fonction de leur practice 
    tuData.forEach(data =>{
      const TuDatas = {...data}
      if(!TuDatas.dataValues.name){
        throw new Error ('data name is null')
      }

      CollaboratorsInfos.forEach(Collab => {
        if(!Collab.name || Collab.practice){
          throw new Error ('Collab name or practice is null')
        }
        if (TuDatas.dataValues.name == Collab.name){
          if(Collab.practice == ('APPLICATION_ENGINEERING_SERVICES' || 'CLOUD_ADVISORY' || 'COMPLEX_SI_AND_ARCHITECTURE' || 'CORE_AMS' || 'DEVSECOPS_&_IT_AUTOMATION')){ //Hybrides Cloud Service
            HybridesCloudService.add(TuDatas.dataValues)
            if (Collab.practice == ('CORE_AMS' || 'DEVSECOPS_&_IT_AUTOMATION')){
              HybridesCloudManagement.add(TuDatas.dataValues)
            }else {
              HybridesCloudTransformation.add(TuDatas.dataValues)
            }
          } else if (Collab.practice == ('CX_&_COMMERCE_/_ADOBE' || 'EXPERIENCE_DESIGN_&_MOBILE' || 'ENTERPRISE_STRATEGY' || 'SALESFORCE' || 'ORACLE' || 'WORKDAY_FINANCE' || 'SAP' || 'BLOCKCHAIN' || 'DATA_&_ANALYTICS' || 'IA_&_AUTOMATION' || 'IOT_&_ASSET_MANAGEMENT' || 'MICROSOFT' || 'CAPITAL_MARKET,_RISK_&_COMPLIANCE' || 'INSURANCE' || 'RETAIL_BANKING' || 'ENTERPRISE_CHANGE_&_TRANSFORMATION' || 'WORKDAY_&_HCM_CLOUD')){ //BTS#
            BTS.add(TuDatas.dataValues)
            if (Collab.practice == ('CX_&_COMMERCE_/_ADOBE' || 'EXPERIENCE_DESIGN_&_MOBILE')){
              CustomerTransformation.add(TuDatas.dataValues)
            }else if (Collab.practice == ('ENTERPRISE_STRATEGY')){
              EntrepriseStartegy.add(TuDatas.dataValues)
            }else if (Collab.practice == ('SALESFORCE')){
              SalesForce.add(TuDatas.dataValues)
            }else if (Collab.practice == ('ORACLE' || 'WORKDAY_FINANCE' || 'SAP')) {
              FinanceChainTransformation.add(TuDatas.dataValues)
            }else if (Collab.practice == ('WORKDAY_&_HCM_CLOUD' || 'ENTERPRISE_CHANGE_&_TRANSFORMATION')){
              TalentTransformation.add(TuDatas.dataValues)
            }else if (Collab.practice == ('CORE_AMS' || 'DEVSECOPS_&_IT_AUTOMATION' || 'CORE_AMS' || 'DEVSECOPS_&_IT_AUTOMATION')){
              IndusrtyTransformation.add(TuDatas.dataValues)
            }else {
              DataTechnologyTransformation.add(TuDatas.dataValues)
            }
          } else {
            throw new Error (Collab.name + ' n\'apartient a aucune practice connu')
          }
        }
      })
    })

    // Chargement du modèle Excel existant
    let workbook;
    try {
      workbook = loadWorkbook();
    } catch (error) {
      throw new Error ('Impossible de charger le template : ' + error)
    }

    // Récupération de la feuille de calcul 'Synthèse'
    let sheetName1;
    let sheetName2;
    let syntheseSheet;
    try {
      sheetName1 = 'Synthèse';
      syntheseSheet = getWorksheet(workbook, sheetName1);
    } catch (error) {
      throw new Error ('Impossible de récupérer de la feuille de calcul "Synthèse" : ' + error)
    }

    // Récupération de la feuille de calcul 'Détails'
    let detailSheet;
    try {
      sheetName2 = 'Détail';
      detailSheet = getWorksheet(workbook, sheetName2)
    } catch (error) {
      throw new Error ('Impossible de récupérer de la feuille de calcul "Détail" : ' + error)
    }

    ExcelSynthLine.add(HybridesCloudService, HybridesCloudManagement, HybridesCloudTransformation, BTS, CustomerTransformation, EntrepriseStartegy, SalesForce, FinanceChainTransformation, DataTechnologyTransformation, IndusrtyTransformation, TalentTransformation)

    ExcelSynthLine.forEach(platform => {
      // Mapping des Colones de la fiche synthèse
      let lineName
      const mappedData = platform.map(item => {
        lineName = item
        return {
          //__TU Projection_____________________________________________________________________________________________________________________________        
          'D': item.CtlJourTuTotalIbmI,
          'E': item.TuTotaFermeFact, 
          'F': item.TuTotalFermeFactChargeable, 
          'G': item.TuTotalFermeFactChargeableProductive, 
          'H': item.TuTotalFactFermePrevisionel, 
          'I': item.TuTotalFactFermePrevisionelRdMp, 
          'J': item.JoursTuTotal, 
          'K': item.FacturableFermeProjection, 
          'L': item.ChargeableHoursProjection, 
          'M': item.ProductiveHoursProjection, 
          'N': item.FacturablePrévisonnel, 
          'O': item.FacturableRdMp, 
          'P': item.CongésRttGérable, 
          'Q': item.CongésRttNonGérable, 
          'R': item.JoursFériés, 
          'S': item.MaladiesMaternité, 
          'T': item.FormationFerme, 
          'U': item.FormationPrévisionnelle, 
          'V': item.ActivitésInternesFermes, 
          'W': item.ActivitésInternesPrévisionnelles, 
          'X': item.AutresNonFacturables, 
          'Y': item.DispoPrév, 
          'Z': item.TotalNonFacturable, 
          //__TU Reel _____________________________________________________________________________________________________________________________        
          'AA': item.TuQtdFacturable, 
          'AB': item.TuQtdFacturableChargeable, 
          'AC': item.TuQtdFacturableChargeableProductive, 
          'AD': item.JoursTu, 
          'AE': item.Facturable, 
          'AF': item.ChargeableHours, 
          'AG': item.ProductiveHours, 
          'AH': item.CongésRttGérable, 
          'AI': item.CongésRttNonGérable, 
          'AJ': item.JoursFériés, 
          'AK': item.MaladiesMaternité, 
          'AL': item.Formation, 
          'AM': item.AutresNonFacturables, 
          'AN': item.disponibilité, 
          'AO': item.TotalNonFacturable, 
          //__TU prévisionnel _____________________________________________________________________________________________________________________________        
          'AP': item.TuPrévisionelFactFerme, 
          'AQ': item.TuPrévisionelFactChargeableFerme, 
          'AR': item.TuPrévisionelFactChargeableProductiviteFerme, 
          'AS': item.TuPrévisionelFactFermePrévisionnel, 
          'AT': item.TuPrévisionelFactFermePrévisionnelRdMp, 
          'AU': item.JoursTu, 
          'AV': item.FacturableFerme, 
          'AW': item.ChargeableHoursPrev, 
          'AX': item.ProductiveHoursPrev, 
          'AY': item.FacturablePrévisonnel, 
          'AZ': item.FacturableRdMp, 
          'BA': item.CongésRttValides, 
          'BB': item.CongésRttPrévisionnel, 
          'BC': item.JoursFériés, 
          'BD': item.MaladiesMaternité, 
          'BE': item.FormationFerme, 
          'BF': item.FormationPrévisionnelle, 
          'BG': item.ActivitésInternesFermes, 
          'BH': item.ActivitésInternesPrévisionnelles, 
          'BI': item.AutresNonFacturables, 
          'BJ': item.DispoPrév, 
          'BK': item.TotalNonFacturable, 
          //__ Autres _____________________________________________________________________________________________________________________________        
          'BM': item.TargetBillable, 
          'BO': item.PvBillable, 
          'BQ': item.DeltaToTarget, 
          'BT': item.DeltaPrevPv 
        };
      });
      // Conversion des données en format Excel dela fiche synthèse
      let data;
      let newData;
      try {
        data = XLSX.utils.sheet_to_json(syntheseSheet);
        newData = [...data, ...mappedData];
        if (lineName == 'CustomerTransformation'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D10' }); //commence a remplir a partir de la case D10
        }else if (lineName == 'EntrepriseStartegy'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D14' }); //commence a remplir a partir de la case D14
        }else if (lineName == 'SalesForce'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D16' }); //commence a remplir a partir de la case D16
        }else if (lineName == 'FinanceChainTransformation'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D17' }); //commence a remplir a partir de la case D17
        } else if (lineName == 'DataTechnologyTransformation'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D22' }); //commence a remplir a partir de la case D22
        }else if (lineName == 'IndusrtyTransformation'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D28' }); //commence a remplir a partir de la case D28
        }else if (lineName == 'TalentTransformation'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D33' }); //commence a remplir a partir de la case D33
        }else if (lineName == 'HybridesCloudService'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D36' }); //commence a remplir a partir de la case D36
        }else if (lineName == 'HybridesCloudManagement'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D37' }); //commence a remplir a partir de la case D37
        }else if (lineName == 'HybridesCloudTransformation'){
          XLSX.utils.sheet_add_json(syntheseSheet, newData, { skipHeader: true, origin: 'D40' }); //commence a remplir a partir de la case D40
        }
      } catch (error) {
        throw new Error ('Impossible de convertir les données en format Excel pour la feuille synthèse : ' + error)
      }
    })

    // Mapping des Colones de la fiche détail
    const detailMappedData = Collaborator.map(item => {
      return {
        //__Collaborateur_____________________________________________________________________________________________________________________________        
        'B': item.inIAS,
        'C': item.outIAS,
        'D': item.inTuCalculation,
        'E': item.outTuCalculation, 
        'F': item.dispatched, 
        'G': item.inTu, 
        // valeur a revoir inTu
        'H': item.partialTime, 
        'I': item.percentPartialTime, 
        'J': item.dayPartialTime, 
        'K': item.name, 
        'L': item.id, 
        //
        'M': item.trCost, 
        'N': item.site, 
        'O': item.serviceLine, 
        'P': item.practice, 
        'Q': item.sat, 
        'R': item.practice, 
        'S': item.sat, 
        'T': item.responsableCentredeService, 
        //
        'U': item.trCost, 
        'V': item.site, 
        'W': item.serviceLine, 
        'X': item.practice, 
        'Y': item.sat, 
        'Z': item.responsableCentredeService, 
        //__TU Reel _____________________________________________________________________________________________________________________________        
        'AA': item.autresCaractéristiques, 
        'AB': item.regroupementAutre, 
        'AC': item.localisationPerso, 
        'AD': item.JoursTu, 
        'AE': item.Facturable, 
        'AF': item.ChargeableHours, 
        'AG': item.ProductiveHours, 
        'AH': item.CongésRttGérable, 
        'AI': item.CongésRttNonGérable, 
        'AJ': item.JoursFériés, 
        'AK': item.MaladiesMaternité, 
        'AL': item.Formation, 
        'AM': item.AutresNonFacturables, 
        'AN': item.disponibilité, 
        'AO': item.TotalNonFacturable, 
        'AP': item.TuPrévisionelFactFerme, 
        'AQ': item.TuPrévisionelFactChargeableFerme, 
        'AR': item.TuPrévisionelFactChargeableProductiviteFerme, 
        'AS': item.TuPrévisionelFactFermePrévisionnel, 
        'AT': item.TuPrévisionelFactFermePrévisionnelRdMp, 
        'AU': item.JoursTu, 
        'AV': item.FacturableFerme, 
        'AW': item.ChargeableHoursPrev, 
        'AX': item.ProductiveHoursPrev, 
        'AY': item.FacturablePrévisonnel, 
        'AZ': item.FacturableRdMp, 
        'BA': item.CongésRttValides, 
        'BB': item.regroupementAutre, 
        'BC': item.localisationPerso, 
        'BD': item.JoursTu, 
        'BE': item.Facturable, 
        'BF': item.ChargeableHours, 
        'BG': item.ProductiveHours, 
        'BH': item.CongésRttGérable, 
        'BI': item.CongésRttNonGérable, 
        'BJ': item.JoursFériés, 
        'BK': item.MaladiesMaternité, 
        'BL': item.Formation, 
        'BM': item.AutresNonFacturables, 
        'BN': item.disponibilité, 
        'BO': item.TotalNonFacturable, 
        'BP': item.TuPrévisionelFactFerme, 
        'BQ': item.TuPrévisionelFactChargeableFerme, 
        'BR': item.TuPrévisionelFactChargeableProductiviteFerme, 
        'BS': item.TuPrévisionelFactFermePrévisionnel, 
        'BT': item.TuPrévisionelFactFermePrévisionnelRdMp, 
        'BU': item.JoursTu, 
        'BV': item.FacturableFerme, 
        'BW': item.ChargeableHoursPrev, 
        'BX': item.ProductiveHoursPrev, 
        'BY': item.FacturablePrévisonnel, 
        'BZ': item.FacturableRdMp, 
        'CA': item.CongésRttValides, 
        'CB': item.CongésRttPrévisionnel, 
        'CC': item.JoursFériés, 
        'CD': item.MaladiesMaternité, 
        'CE': item.FormationFerme, 
        'CF': item.FormationPrévisionnelle, 
        'CG': item.ActivitésInternesFermes, 
        'CH': item.ActivitésInternesPrévisionnelles, 
        'CI': item.AutresNonFacturables, 
        'CJ': item.DispoPrév, 
        'CK': item.TotalNonFacturable, 
        'CL': '',
        'CM': item.TotalJourQ, 
        'CN': item.TotalJourQE,
        'CO': item.JourFacturableQE, 
        'CP': item.JourTuQTD,
        'CQ': item.JourFacturableQTD, 
        'CR': item.JourTuToGo,
        'CS': item.JourFacturableToGo,
        'CT': item.Fte,
        'CU': item.Plateform,
        'CV': item.Tranche,
        'CW': item.Bench,
        'CX': item.MajorationDispo,
        'CY': item.MajorationPrev,
        'CZ': item.DispoPrev,
        'DA': item.FacturablePrévisonnel,
        'DB': item.FacturableFerme,
      };
    });
    
    // Conversion des données en format Excel dela fiche détail
    let detailData;
    let newDetailData;
    try {
      detailData = XLSX.utils.sheet_to_json(detailSheet);
      newDetailData = [...detailData, ...detailMappedData];
      XLSX.utils.sheet_add_json(detailSheet, newDetailData, { skipHeader: true, origin: 'A5' }); //commence a remplir a partir de la case A5
    } catch (error) {
      throw new Error ('Impossible de convertir les données en format Excel pour la feuille détail : ' + error)
    }
    

    // Enregistrer le classeur sous un nouveau nom
    const outputFilePath = path.join(__dirname, '../assets/tu_data.xlsx');
    XLSX.writeFile(workbook, outputFilePath);

    // Envoi du fichier Excel en réponse à la requête du client React
    res.download(outputFilePath, 'tu_data.xlsx', (err) => {
      if (err) {
        console.error(err);
      } else {
        // Suppression du fichier Excel après envoi
        fs.unlink(outputFilePath, (err) => {
          if (err) {
            console.error(err);
          } else {
            console.log('File deleted successfully');
          }
        });
      } 
    });
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Service Error');
  }
}

module.exports = { exportTuFile };
