const fs = require('fs');

const classes = `
    classDef root fill:#1e1e2e,stroke:#89b4fa,stroke-width:2px,color:#cdd6f4,font-weight:bold,font-size:16px
    classDef choice fill:#313244,stroke:#cba6f7,stroke-width:1px,color:#e0dff0
    classDef check fill:#181825,stroke:#a6e3a1,stroke-width:2px,color:#a6e3a1,stroke-dasharray: 4 4
`;

let output = '%% おかゆにゅ～～む！R コンプリート用（全パターン網羅）フローチャート\ngraph LR\n' + classes + '\n';
let counter = 1;

function makePaths(phaseObj) {
    let out = '';
    out += `    subgraph ${phaseObj.id} [${phaseObj.title} 全${phaseObj.total}パターン]\n        direction LR\n        ${phaseObj.id}_root(${phaseObj.rootName}):::root\n\n`;
    
    function generateCombinations(arrays) {
        if (arrays.length === 0) return [[]];
        let result = [];
        let rest = generateCombinations(arrays.slice(1));
        for (let item of arrays[0]) {
            for (let r of rest) {
                result.push([item, ...r]);
            }
        }
        return result;
    }

    let combinations = generateCombinations(phaseObj.levels);

    combinations.forEach((combo, idx) => {
        let parent = `${phaseObj.id}_root`;
        for(let i=0; i<combo.length; i++) {
            let itemName = combo[i];
            let currentId = `${phaseObj.id}_${idx}_${i}`;
            
            if(i === combo.length - 1) {
                // Leaf
                let leafID = `L_${counter++}`;
                out += `        ${parent} --> ${leafID}["[ メモ：　　　 ] ${itemName}"]:::check\n`;
            } else {
                // Determine if this choice node already exists to avoid duplicates in the graph
                // Since this is a check matrix, we WANT duplicates visually, or do we?
                // Visual branching is better.
            }
        }
    });
    
    // Wait, generating full trees is better done recursively to merge prefixes.
    return out;
}

function traverseAndMerge(phaseObj) {
    let out = '';
    out += `    subgraph ${phaseObj.id} [${phaseObj.title} 全${phaseObj.total}パターン]\n        direction LR\n        ${phaseObj.id}_root(${phaseObj.rootName}):::root\n\n`;
    
    function traverse(nodeID, currentDepth) {
        if(currentDepth >= phaseObj.levels.length) {
            let leafID = `L_${counter++}`;
            out += `        ${nodeID} --> ${leafID}["[ 回収メモ： 未記入 ]"]:::check\n`;
            return;
        }
        
        let items = phaseObj.levels[currentDepth];
        for(let i=0; i<items.length; i++) {
            let itemName = items[i];
            let childID = `${nodeID}_${i}`;
            
            if (currentDepth === phaseObj.levels.length - 1){
                let leafID = `L_${counter++}`;
                out += `        ${nodeID} --> ${leafID}["[ 回収：　　　 ] ${itemName}"]:::check\n`;
            } else {
                out += `        ${nodeID} --> ${childID}[${itemName}]:::choice\n`;
                traverse(childID, currentDepth + 1);
            }
        }
    }
    
    traverse(`${phaseObj.id}_root`, 0);
    out += `    end\n\n`;
    return out;
}

const phases = [
    { id: 'P1', title: '第1フェーズ：にゃふにゃふT', rootName: 'アイテム', total: 18,
      levels: [['布団', 'シャツ', '野球拳'], ['おでこ', 'にのうで'], ['紅葉狩り【きのこ】', '紅葉狩り【どんぐり】', 'リンゴ狩り']] },
    { id: 'P2', title: '第2フェーズ：クイズ＆探索', rootName: '行き先', total: 6,
      levels: [['酸ヶ湯', '青森市', '八戸市'], ['たまご味噌'], ['みみ', 'ふともも']] },
    { id: 'P3', title: '第3フェーズ：ASMR', rootName: 'プラン', total: 3,
      levels: [['津軽鉄道【車内】', '津軽鉄道【車外】', '市場']] },
    { id: 'P4', title: '第4フェーズ：にゃふにゃふ中盤', rootName: 'スキンシップ', total: 4,
      levels: [['抱き寄せる', '飲み物差し入れ'], ['あご', '手']] }, 
    { id: 'P5', title: '第5フェーズ：おかゆなでなで', rootName: 'おみくじ', total: 12,
      levels: [['左のおみくじ', '右のおみくじ'], ['はな', 'おなか'], ['氷爆【五目並べ】', '氷爆【TVゲーム】', '水族館']] },
    { id: 'P6', title: '第6フェーズ：にゃふにゃふ後半1', rootName: 'アイテム', total: 12,
      levels: [['手振り鉦', 'ねぶた絵'], ['頭', '尻尾'], ['樹氷【見に行く】', '樹氷【一休み】', 'ウィンタースポーツ']] },
    { id: 'P7_A', title: '第7フェーズ：撮影タイムA', rootName: 'アイテム', total: 12,
      levels: [['バレンタイン', 'おにぎり'], ['おなか', '頬'], ['ピース【顔】', 'ネコ【尻尾】', 'ハート【胸】']] },
    { id: 'P7_B', title: '第7フェーズ：お布団スキンシップ', rootName: 'アイテム', total: 4,
      levels: [['お布団', '妹モノの本'], ['ふともも', '鎖骨']] },
    { id: 'P7_C', title: '第7フェーズ：製菓スキンシップ', rootName: 'アイテム', total: 4,
      levels: [['製菓キット', 'チョコの箱'], ['あご', '肩']] },
    { id: 'P8', title: '第8フェーズ：クライマックス', rootName: 'お風呂の味', total: 6,
      levels: [['抹茶', 'チョコ', 'バニラ'], ['腹', 'ペロ']] }
];

phases.forEach(p => output += traverseAndMerge(p));

fs.writeFileSync('C:/Users/motec/Desktop/WARP/WARP-002-034/Claude-Code-Genereter/マインドマップ/おかゆにゅ～～む_コンプ用.mmd', output, 'utf8');
console.log('Saved to おかゆにゅ～～む_コンプ用.mmd');
