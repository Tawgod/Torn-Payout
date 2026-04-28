const SETTINGS = {
  configSheet: "Config",
  apiKeyCell: "B1",           
  backendIdCell: "B2",        
  archiveIdCell: "B3",        
  targetFactionCell: "B4",    
  warStartCell: "B5",         
  warIdCell: "B6",            
  
  // NEW: Tell the script where your new dropdowns live on the Dashboard
  dashOutcomeCell: "C10", 
  dashTermedCell: "C15",

  rdSheet: "RD",
  dashboardSheet: "Dashboard",
  rosterSheet: "Roster", 
  payoutSheet: "Payouts",
  bountySheet: "Bounties",   
  finalSheet: "Final Payout", 
  historySheet: "History", 

  rdHeaders: [
    "Attack ID", "Start Time", "End Time", "Attacker ID", "Attacker Name", "Attacker Faction", "Defender ID", "Defender Name", "Defender Faction", "Result", "Respect", "Fair Fight", "War Bonus", "Retaliation", "Group Attack", "Overseas", "Chain Bonus", "Warlord Bonus"
  ],
  rosterHeaders: [ 
    "Member ID", "Name", "Level", "Position", "Days in Faction", "Status", "Last Action"
  ],

  // UPDATED: Added Outcome and Termed right after the Enemy Faction
  historyHeaders: [ 
    "Date Archived", "War ID", "Enemy Faction", "Outcome", "Termed", 
    "Total Revenue", "Total Costs", "Net Profit", "Total Payout",
    "Wt: War Hits", "Wt: War Assists", "Wt: War Losses", "Wt: War Interrupts", 
    "Wt: Outside/Chain", "Wt: Chain Saves", "Wt: Retaliations", 
    "Wt: Net Respect", "Wt: War Abroad", "Wt: War Score" 
  ],
  bountyHeaders: [           
    "Date Logged", "Placed By", "Target", "Bounty Amount", "Refund Amount", "Status", "Notes"
  ],
  finalHeaders: [            
    "Member ID", "Name", "War Payout ($)", "Bounty Refunds ($)", "Misc. Adjustments ($)", "Total Final Payout ($)", "One-Click Pay Link", "Payment Status" 
  ]
};