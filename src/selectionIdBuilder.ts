module powerbi.extensibility.visual {
  // powerbi.visuals
  import ISelectionId = powerbi.visuals.ISelectionId;
  // powerbi.extensibility.visual
  import IVisualHost = powerbi.extensibility.visual.IVisualHost;

  interface CategoryIdentityIndex {
    categoryIndex: number;
    identityIndex: number;
  }

  export class CalendarSelectionIdBuilder {

    private static DefaultCategoryIndex: number = 0;
    private visualHost: IVisualHost;
    private categories: DataViewCategoryColumn[];

    constructor(IVisualHost: IVisualHost, categories: DataViewCategoryColumn[]) {
      this.visualHost = IVisualHost;
      this.categories = categories || [];
    }

    private getIdentityById(index: number): CategoryIdentityIndex {
      let categoryIndex: number = CalendarSelectionIdBuilder.DefaultCategoryIndex;
      let identityIndex: number = index;

      for (let length: number = this.categories.length; categoryIndex < length; categoryIndex++) {
        let amountOfIdentities: number = this.categories[categoryIndex].identity.length;
        if (identityIndex > amountOfIdentities - 1) {
          identityIndex -= amountOfIdentities;
        } else {
          break;
        }
      }

      return {
        categoryIndex,
        identityIndex
      };
    }

    public createSelectionId(index: number): ISelectionId {
      let categoryIdentityIndex: CategoryIdentityIndex = this.getIdentityById(index);

      return this.visualHost.createSelectionIdBuilder()
        .withCategory(this.categories[categoryIdentityIndex.categoryIndex], categoryIdentityIndex.identityIndex)
        .createSelectionId();
    }
  }
}
