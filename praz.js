const cnst = 5;
var entry = 'A';
var i = 0;

function main() {
    console.clear();
    console.log('--- START ---\n');

    var sum = 0;
    var yy1 = i + 10;
    var y2 = 0;
    //y2 = y1 * 2;
    
    function start(){
        console.log('--- start ');

        var p1 = 250;
        var p2 = 350;
        y2 = yy1 * 3; // y2 gets really modified

        return {    // you dont even need to return y2
            one: p1,
            two: p2
        };
    }

    function run(){
        console.log('--- run ');
        var yeah = start().one + yy1;
        console.log('output run = ' + yeah);
    }

    function finish(p1, p2){
        console.log('--- finish ');

        sum = p1 + p2;
        i = yy1 + y2; // y2 gets really recognized
        
        return {
            one: sum,
            two: i
        };
    }

    run();
    entry = finish(start().one, start().two);
    console.log(entry.one); // == finish().one
    console.log(entry.two); // == finish().two
    
    // while(i < cnst){

    //     function pro13(){ // i
    //         try{
    //             console.log('--- const pro13 ' + cnst);
    //         }
    //         catch{
    //             console.log('\n--- pro13 XX ---\n');
    //         }
    //         var p13 = 11;
    //         return p13 + i;
    //     }

    //     var remember13 = pro13(); // i
    //     console.log('--- console pro13 ' + remember13);

    //     function pro16(){
    //         try{
    //             console.log('--- pro16 ' + remember13 + entry);
    //         }
    //         catch{
    //             console.log('\n--- pro16 XX ---\n');
    //         }
    //         var p16 = 22;
    //         return p16 + entry;
    //     }

    //     console.log('--- console ' + pro16() + '\n');
    //     i++;
    // }
}

main();