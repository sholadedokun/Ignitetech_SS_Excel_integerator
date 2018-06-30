exports.sheetsToEdit = [
  {
    graphRef: 'Joe Chart',
    smartSheetRef: 8150311270410116,
		/*
         --smartsheet reference--
         -[0]first column
         -[1]first row
         -[2]last column
         -[3]last row

         --graphReference--
         [0]startRange
         [1]lastRage
        */
    rangeValue: [
      {
        smartSheet: [3, 16, 22, 16],
        graph: ['C71', 'V71']
      },
      {
        smartSheet: [7, 19, 11, 19],
        graph: ['G74', 'K74']
      },
      {
        smartSheet: [19, 38, 19, 41],
        graph: ['S92', 'S95']
      },
      {
        smartSheet: [18, 54, 18, 57],
        graph: ['R108', 'R111']
      },
      {
        smartSheet: [15, 5, 16, 5],
        graph: ['V60', 'W60']
      },
      {
        smartSheet: [4, 2, 5, 2],
        graph: ['Y33', 'Z33']
      }
    ]
  },
  {
    graphRef: 'P&L',
    includeSSColumn: 0, // to include the smartsheet column or not during copy
    smartSheetRef: 6602198898501508
  },
  {
    graphRef: 'HC Legacy',
    includeSSColumn: 1,
    rangeEdit: 0,
    smartSheetRef: 1535649317709700
  },
  {
    graphRef: 'HC Ongoing',
    includeSSColumn: 1,
    smartSheetRef: 2098599271131012
  },
  {
    graphRef: 'Non HC',
    includeSSColumn: 0,
    smartSheetRef: 5248700218926980
  },
  {
    graphRef: 'Units Budget',
    includeSSColumn: 1,
    smartSheetRef: 6039248945080196
  },
  {
    graphRef: 'Customer Dashboard',
    includeSSColumn: 1,
    smartSheetRef: 3787449131394948
  },
  {
    graphRef: 'MR Budget',
    includeSSColumn: 1,
    smartSheetRef: 8291048758765444
  }

];
