(function(){
    function create2DArray(N,M,defaultValue=''){
        let arr=new Array(N);
        for(let i=0;i<N;i++){
            if(typeof defaultValue=="object"){
                arr[i]=new Array(M).fill({...defaultValue});//to give a new object each time so as to avoid reference sharing
            }else{
                arr[i]=new Array(M).fill(defaultValue);
            }
        }
        return arr;
    }
    function numbertoColumnName(num){
        let columnName='';
        while(num>0){
            let rem=(num-1)%26;
            columnName=String.fromCharCode(65+rem)+columnName;
            num=Math.floor((num-1)/26);
        }
        return columnName;
    }

    var textInput;
    const rowCount=100;
    const colCount=100;
    let mode="NORMAL";//normal or search


    let selectedCell={
        row:1,col:1 //0th index reserved for row names and column names
    }
    let selectedCellRange={
        startRow:selectedCell.row,
        endRow:selectedCell.row,
        startCol:selectedCell.col,
        endCol:selectedCell.col
    }//range selected by user
    let spreadsheetData=create2DArray(rowCount,colCount,'');
    let cellProperties=create2DArray(rowCount,colCount,{
        textAlign:"left"   //left-aligned by default
    });

    let formulaInCell={};
    let execDependentOnCell={}
     
    let canvas;
    let ctx;
    let rowHeight=new Array(rowCount).fill(25);//size of each row
    let colWidth=new Array(colCount).fill(100);//size of each column

    for(let r=0;r<rowCount;r++){
        for(let c=0;c<colCount;c++){
            if(r==0&&c>0){
                spreadsheetData[r][c]=numbertoColumnName(c);
                cellProperties[r][c]['textAlign']='center';
            }
            else if(c==0&&r>0){
                spreadsheetData[r][c]=r.toString();
                cellProperties[r][c]['textAlign']='center';
            }
        }
    }

    var DrawFunctions=(function(){
        function getCellPosition(row,col){
            let y=0;
            for(let r=0;r<row;r++){
                y+=rowHeight[r];
            }//to get the y-coordinate of the cell
            let x=0;
            for(let c=0;c<col;c++){
                x+=colWidth[c];
            }//to get the x-coordinate of the cell
            return {x,y};
        }

        function drawSingleCell(row,col){
            const props=cellProperties[row][col];
            const pos=getCellPosition(row,col);
    
            ctx.strokeRect(pos.x,pos.y,colWidth[col],rowHeight[row]);

            const text=spreadsheetData[row][col];

            ctx.font='12px sans-serif';
            ctx.textAlign='left';
            ctx.textBaseline='middle';

            if('textAlign' in props){
                ctx.textAlign=props['textAlign'];
            }
            const paddingX=5;
            var textX;
            switch(ctx.textAlign){
                case 'left':
                    textX=pos.x+paddingX;
                    break;
                case 'right':
                    textX=pos.x+colWidth[col]-paddingX;
                    break;
                case 'center':
                    textX=pos.x+colWidth[col]/2;
                    break;
            }

            const textY=pos.y+rowHeight[row]/2;

            ctx.fillText(text,textX,textY);
        }
        function drawCells(){
            for(let i=0;i<rowCount;i++){
                for(let j=0;j<colCount;j++){
                    drawSingleCell(i,j)
                }
            }
        }
        function drawSelectedCellBorder(){
            const pos=getCellPosition(selectedCell.row,selectedCell.col);
            const borderDiv=document.getElementById("selectedCellBorder");
            borderDiv.style.display="block";
            borderDiv.style.left=`${pos.x+1}px`;//+1 to avoid overlapping with the cell border
            borderDiv.style.top=`${pos.y+1}px`;
            borderDiv.style.width=`${colWidth[selectedCell.col]-1}px`;
            borderDiv.style.height=`${rowHeight[selectedCell.row]-1}px`;
            borderDiv.style.border="3px solid blue";
        }

        function drawSelectedCellRangeBorder(){
            const selectedRangeDiv=document.getElementById("selectedRange");
            
            let rangeStartCol=Math.min(selectedCellRange.startCol,selectedCellRange.endCol);
            let rangeStartRow=Math.min(selectedCellRange.startRow,selectedCellRange.endRow);
            const pos=getCellPosition(rangeStartRow,rangeStartCol);
            selectedRangeDiv.style.display="block";
            selectedRangeDiv.style.left=`${pos.x+1}px`;
            selectedRangeDiv.style.top=`${pos.y+1}px`;
            let rangeWidth=0;
            let rangeHeight=0;
            let startRow=Math.min(selectedCellRange.startRow,selectedCellRange.endRow);
            let endRow=Math.max(selectedCellRange.startRow,selectedCellRange.endRow);

            for(let r=startRow;r<=endRow;r++){
                rangeHeight+=rowHeight[r];
            }
            let startCol=Math.min(selectedCellRange.startCol,selectedCellRange.endCol);
            let endCol=Math.max(selectedCellRange.startCol,selectedCellRange.endCol);
            for(let c=startCol;c<=endCol;c++){
                rangeWidth+=colWidth[c];
            }
            selectedRangeDiv.style.width=`${rangeWidth-1}px`;
            selectedRangeDiv.style.height=`${rangeHeight-1}px`;
            if(selectedCellRange.startRow!=selectedCellRange.endRow||selectedCellRange.startCol!=selectedCellRange.endCol){
                selectedRangeDiv.style.border="1px solid rgba(0,0,255,0.8)";
                selectedRangeDiv.style.backgroundColor="rgba(0,0,255,0.3)";
            }else{
                selectedRangeDiv.style.backgroundColor="rgba(0,0,255,0)";
                selectedRangeDiv.style.backgroundColor="rgba(0,0,255,0)";
            }
        }

        function drawGrid(){
            ctx.clearRect(0,0,canvas.width,canvas.height);
            ctx.strokeStyle='rgba(0, 0, 0, 0.2)';
            drawCells();
            drawSelectedCellBorder();
            drawSelectedCellRangeBorder();
        }
        return {drawGrid,getCellPosition};
    })()

    

    document.addEventListener("DOMContentLoaded",function(){
        canvas=document.getElementById("spreadsheet");
        textInput=document.getElementById("textInput");
        ctx=canvas.getContext("2d");
        const canvasWidth=1200;
        const canvasHeight=1700;
        const ratio=window.devicePixelRatio;
        canvas.width=canvasWidth*ratio;
        canvas.height=canvasHeight*ratio;
        canvas.style.width=canvasWidth+"px";
        canvas.style.height=canvasHeight+"px";
        ctx.scale(ratio,ratio);

        DrawFunctions.drawGrid()//this prevents exposure of helper functions to global scope ie 
    })
    function checkSelectedCellShift(key){
        const r=selectedCell.row;
        const c=selectedCell.col;
        if(key=="ArrowLeft"){
            return c>1;
        }
        if(key=="ArrowRight"){
            return c<colCount-1;
        }
        if(key=="ArrowUp"){
            return r>1;
        }
        if(key=="ArrowDown"){
            return r<rowCount-1;
        }
    }

    function resetSelectedCellRange(){
        selectedCellRange.startRow=selectedCell.row;
        selectedCellRange.endRow=selectedCell.row;
        selectedCellRange.startCol=selectedCell.col;
        selectedCellRange.endCol=selectedCell.col;
    }

    function isFormula(cell){
        return cell.startsWith("=");
    }

    function convertCellNametoIndex(cellName){
        let columnLetters=cellName.match(/[A-Z]+/)[0];
        let rowNumber=parseInt(cellName.match(/\d+/)[0],10);

        let col=0;
        for(let i=0;i<columnLetters.length;i++){
            col=col*26+(columnLetters.charCodeAt(i)-'A'.charCodeAt(0)+1);
        }
        return {
            row:rowNumber,
            col:col
        }
    }

    function convertIndexToCellName(row,col){
        return numbertoColumnName(col)+row.toString();
    }

    function run_avg_func(paramString,opCellName){
        const sum=run_sum_func(paramString,opCellName);
        if(isFinite(sum)){
            const cellRange=paramString;
            const cellNameRegex=/[A-Z]+\d+/g;
            const cellNames=cellRange.match(cellNameRegex);
            if(cellNames.length!=2){
                return "NaN";
            }
            
            const c1=convertCellNametoIndex(cellNames[0]);
            const c2=convertCellNametoIndex(cellNames[1]);
            let numCells=(Math.abs(c1.row-c2.row)+1)*(Math.abs(c1.col-c2.col)+1);
            return sum/numCells;
            
        }
        return "NaN";
    }

    function run_sum_func(paramString,opCellName){
        const cellRange=paramString;
        const cellNameRegex=/[A-Z]+\d+/g;
        const cellNames=cellRange.match(cellNameRegex);
        if(cellNames.length!=2){
            return "NaN";
        }
        let sum=0;
        const c1=convertCellNametoIndex(cellNames[0]);
        const c2=convertCellNametoIndex(cellNames[1]);

        for(let r=Math.min(c1.row,c2.row);r<=Math.max(c1.row,c2.row);r++){
            for(let c=Math.min(c1.col,c2.col);c<=Math.max(c1.col,c2.col);c++){
                //add dependecies
                const cName=convertIndexToCellName(r,c);
                if(!(cName in execDependentOnCell)){
                    execDependentOnCell[cName]=[];
                }
                if(!execDependentOnCell[cName].includes(opCellName)){
                    execDependentOnCell[cName].push(opCellName);
                }

                const v=spreadsheetData[r][c];
                if(isFinite(v)){
                    sum+=parseFloat(v);
                }else{
                    return "NaN";
                }
            }
        }
        return sum;
    }

    function run_min_func(paramString,opCellName){
        const result=run_min_max_func(paramString,opCellName);
        if(typeof result=="object"){
            return result.min;
        }
        return "NaN";
    }
    function run_max_func(paramString,opCellName){
        const result=run_min_max_func(paramString,opCellName);
        if(typeof result=="object"){
            return result.max;
        }
        return "NaN";
    }


    function run_min_max_func(paramString,opCellName){
        const cellRange=paramString;
        const cellNameRegex=/[A-Z]+\d+/g;
        const cellNames=cellRange.match(cellNameRegex);
        if(cellNames.length!=2){
            return "NaN";
        }
        let min=10000;
        let max=-10000;
        const c1=convertCellNametoIndex(cellNames[0]);
        const c2=convertCellNametoIndex(cellNames[1]);

        for(let r=Math.min(c1.row,c2.row);r<=Math.max(c1.row,c2.row);r++){
            for(let c=Math.min(c1.col,c2.col);c<=Math.max(c1.col,c2.col);c++){
                //add dependecies
                const cName=convertIndexToCellName(r,c);
                if(!(cName in execDependentOnCell)){
                    execDependentOnCell[cName]=[];
                }
                if(!execDependentOnCell[cName].includes(opCellName)){
                    execDependentOnCell[cName].push(opCellName);
                }

                const v=spreadsheetData[r][c];
                if(isFinite(v)){
                    const vAsFloat=parseFloat(v);
                    min=Math.min(min,vAsFloat);
                    max=Math.max(max,vAsFloat);
                }else{
                    return "NaN";
                }
            }
        }
        return {min,max};
    }

    const functions=["SUM","AVERAGE","MAX","MIN" ];
    const funcMap={
        "SUM":run_sum_func,
        "AVERAGE":run_avg_func,
        "MAX":run_max_func,
        "MIN":run_min_func
    };

    function isFunctionpresent(expr){
        
        for(let f of functions){
            if(expr.includes(`${f}(`)){
                return true;
            }
        }
        return false;
    }

    function execFunctions(expr,opCellName){
        let extract=false
        let extractFunc=false;
        let parameters=''
        let funcName=''
        for(let ch of expr){
            if(ch=="="){
                extractFunc=true;
                continue;
            }
            if(ch=="("){
                if(extractFunc){
                    extractFunc=false;
                }
                extract=true;
                continue;
            }
            if(ch==")"&&extract){
                extract=false;
                break;
            }
            if(extract){
                parameters+=ch;
                continue;
            }
            if(extractFunc){
                funcName+=ch;
                continue;
            }
        }
        if(funcName in funcMap){
            return funcMap[funcName](paramString=parameters,opCellName=opCellName);
        }
        return "NaN";

    }

    function execute(expr,opCellName){
        if(isFunctionpresent(expr)){
            return execFunctions(expr,opCellName);
        }
        expr=expr.replaceAll("=","")
        const cellNameRegex=/[A-Z]+\d+/g;

        const exprWithVal=expr.replace(
            cellNameRegex,
            (match)=>{
                const rowcol=convertCellNametoIndex(match);

                if(!(match in execDependentOnCell)){
                    execDependentOnCell[match]=[];
                }
                if(!execDependentOnCell[match].includes(opCellName)){
                    execDependentOnCell[match].push(opCellName);
                }
                

                return spreadsheetData[rowcol.row][rowcol.col];
            }
        )
        try{
            return math.evaluate(exprWithVal);
        }
        catch(e){
            console.error('error evaluating $exprWithVal',e);
            return "NaN";
        }
    }

    function reCalculateDependents(cellName){
        if(cellName in execDependentOnCell){
            for(let i=0;i<execDependentOnCell[cellName].length;i++){
                const cellNameToBeExec=execDependentOnCell[cellName][i];
                const cellToBeExec=convertCellNametoIndex(cellNameToBeExec);
                const formula=formulaInCell[cellNameToBeExec];
                const result=execute(formula,cellNameToBeExec);
                spreadsheetData[cellToBeExec.row][cellToBeExec.col]=result;
                updateCellValues(cellNameToBeExec,result);
            }
        }
    }

    function checkCycle(cellName,startCellName,visited=[]){
        //DFS
        if(cellName==startCellName&&visited.length>0){
            return true;
        }
        if(visited.includes(cellName)){
            return false;
        }
        visited.push(cellName);
        if(cellName in execDependentOnCell){
            const dependents=execDependentOnCell[cellName];
            for(let c of dependents){
                if(checkCycle(c,startCellName,visited)){
                    return true;
                }
            }
            return false;
        }
    }

    function updateCellValues(selectedCellName,val){
        //checking for cyclic dependencies
        if(checkCycle(selectedCellName,selectedCellName)){
            alert("Cyclic dependencies detected");
            return;
        }
        const rowcol=convertCellNametoIndex(selectedCellName);
        spreadsheetData[rowcol.row][rowcol.col]=val;
        //go through dependent cells and recalculate the values
        reCalculateDependents(selectedCellName);
    }

    function resetTextInput(){
        if(textInput.style.display=="block"){
            textInput.style.display="none";
            const ipVal=textInput.innerHTML;
            const selectedCellName=convertIndexToCellName(selectedCell.row,selectedCell.col);
            if(isFormula(ipVal)){
                //execute formula
                const result=execute(ipVal,selectedCellName);
                //update formulaInCell and execDependentOnCell
                formulaInCell[selectedCellName]=ipVal;
                updateCellValues(selectedCellName,result);

            }else{
                updateCellValues(selectedCellName,ipVal);
            }
            
            textInput.innerHTML="";
            textInput.blur();
        }
    }

    function handleNormalArrowKeys(key){
        if(checkSelectedCellShift(key)){
            resetTextInput();
            switch(key){
                case "ArrowLeft":
                    selectedCell.col--;
                    break;
                case "ArrowRight":
                    selectedCell.col++;
                    break;
                case "ArrowUp":
                    selectedCell.row--;
                    break;
                case "ArrowDown":
                    selectedCell.row++;
                    break;
            }
            //reset selected cell range
            resetSelectedCellRange();
            
        }
    }
    function handleArrowKeys(key){
        if(mode=="NORMAL"){
            handleNormalArrowKeys(key);
        }
        DrawFunctions.drawGrid();
    }

    function checkSelectedRangeExpansion(key){
        switch(key){
            case "ArrowLeft":
                return selectedCellRange.endCol>1;
                break;
            case "ArrowRight":
                return selectedCellRange.endCol<colCount-1;
                break;
            case "ArrowUp":
                return selectedCellRange.endRow>1;
                break;
            case "ArrowDown":
                return selectedCellRange.endRow<rowCount-1;
                break;
        }
    }

    function handleNormalShiftArrowKeys(key){
        //check for range expansion in direction of key
        if(checkSelectedRangeExpansion(key)){
            switch(key){
                case "ArrowLeft":
                    selectedCellRange.endCol--;
                    break;
                case "ArrowRight":
                    selectedCellRange.endCol++;
                    break;
                case "ArrowUp":
                    selectedCellRange.endRow--;
                    break;
                case "ArrowDown":
                    selectedCellRange.endRow++;
                    break;
            }
        }
    }

    function handleShiftArrowKeys(key){
        if(mode=="NORMAL"){
            handleNormalShiftArrowKeys(key);
        }
        DrawFunctions.drawGrid();
    }

    function displayInput(){
        
        const pos=DrawFunctions.getCellPosition(selectedCell.row,selectedCell.col);
        textInput.style.display="block";
        textInput.style.left=`${pos.x+1}px`;
        textInput.style.top=`${pos.y+1}px`;
        textInput.style.minWidth=`${colWidth[selectedCell.col]-1}px`;
        textInput.style.height=`${rowHeight[selectedCell.row]-1}px`;
        textInput.style.padding="5px"
        textInput.style.backgroundColor="#fff";
        textInput.style.textBaseline="center";
        textInput.style.outline="2px solid blue";
        textInput.value=spreadsheetData[selectedCell.row][selectedCell.col];
        textInput.focus();
    }

    function handleEnterKey(){
        if(textInput.style.display=="none"||textInput.style.display==""){
            displayInput();
        }else if(textInput.style.display=="block"){
            handleArrowKeys("ArrowDown");
        }
    }

    function deleteCellContents(row,col){
        spreadsheetData[row][col]='';
        const selectedCellName=convertIndexToCellName(row,col);
        if(selectedCellName in formulaInCell){
            delete formulaInCell[selectedCellName];
            for(const [k,v] of Object.entries(execDependentOnCell) ){
                if(v.includes(selectedCellName)){
                    const newV=v.filter(function(item){
                        return item!=selectedCellName;
                    })
                    execDependentOnCell[cell]=newV;
                } 
            }
        }
    }

    function handleDeleteKey(){
        deleteCellContents(selectedCell.row,selectedCell.col);
        //delete selected cell range content
        if(selectedCellRange.startRow!=selectedCellRange.endRow&&selectedCellRange.startCol!=selectedCellRange.endCol){
            for(let r=Math.min(selectedCellRange.startRow,selectedCellRange.endRow);r<=Math.max(selectedCellRange.startRow,selectedCellRange.endRow);r++){
                for(let c=Math.min(selectedCellRange.startCol,selectedCellRange.endCol);c<=Math.max(selectedCellRange.startCol,selectedCellRange.endCol);c++){
                    deleteCellContents(r,c);
                }
            }
        }
        DrawFunctions.drawGrid();
    }

    document.addEventListener("keydown",function(event){
        let key=event.key;
        if(event.shiftKey&&['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(key)){
            event.preventDefault();
            handleShiftArrowKeys(key);
        }
        else if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(key)){
            event.preventDefault();
            handleArrowKeys(key);
        } else if(key=="Enter"){
            event.preventDefault();
            handleEnterKey();
        } else if(key.length==1){
            if(textInput.style.display=="none"||textInput.style.display==""){
                displayInput();
            }
        }else if(key=="Backspace"){
            //delete selected cell content and formula
            //delete selected cell range content and formula
            if(textInput.style.display=="none"||textInput.style.display==""){
                event.preventDefault();
                handleDeleteKey();
            }
            
        }
    })

})()