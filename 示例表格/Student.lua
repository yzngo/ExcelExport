--[[
function define:
    GetByIndex( index, prop );
    GetById( id, prop );
    GetIdByIndex( index );
    GetIndexById( id );
    GetCount();

item define:
    ____name                  ____type        ____desc
    index                     int             唯一数字索引
    id                        string          唯一字符串索引
    Age                       int             年龄
    Name                      string          姓名
    Weight                    double          体重
    Skin                      color           肤色
    Language                  group           会说的语言
    Number                    group           数字数组
    Score                     group           小数数组
    Parent                    table           父母信息
姓名等
--]]


--To consider the programmer's habits,replace the 'id' and 'key' in Excel with 'index' and 'id'
local items =
{
    {
        index = 1,
        id = "Stud1",
        Age = 21,
        Name = "XiaoMing",
        Weight = 50.3,
        Skin = 0xFEAC34,
        Language = {"Eng","Chs"},
        Number = {1,2,3,4},
        Score = {1.1,4.5,6.7,4.6},
        Parent = {
            Name = "Haha",
        },
    },

    {
        index = 2,
        id = "Stud2",
        Age = 23,
        Name = "XiaoHua",
        Weight = 56.4,
        Skin = 0xF66666,
        Language = {"Janp","AAA"},
        Number = {35,43,65,63,424},
        Score = {1.1,4.5,6.7,4.6},
        Parent = {
            Name = "AKKK",
        },
    },

}

local indexItems = 
{
    [1] = items[1],
    [2] = items[2],
}

local idItems = 
{
    ["Stud1"] = items[1],
    ["Stud2"] = items[2],
}


local data = { Items = items, IndexItems = indexItems, IdItems = idItems, }

function data:GetByIndex( index, prop )
    local item = self.IndexItems[index];
    if item == nil then
        sGlobal:Print( "Student.lua GetByIndex nil item: "..index );
        return index;
    end
    if prop == nil then
        return item;
    end
    if item[prop] == nil then
        sGlobal:Print( "Student.lua GetByIndex nil prop: "..prop );
        return item;
    end
    return item[prop];
end

function data:GetById( id, prop )
    local item = self.IdItems[id];
    if item == nil then
        sGlobal:Print( "Student.lua GetById nil id: "..id );
        return id;
    end
    if prop == nil then
        return item;
    end
    if item[prop] == nil then
        sGlobal:Print( "Student.lua GetById nil prop: "..prop );
        return item;
    end
    return item[prop];
end

function data:GetIdByIndex( index )
    return data:GetByIndex( index, "id" );
end

function data:GetIndexById( id )
    return data:GetById( id, "index" );
end

function data:GetCount()
    return #self.IndexItems;
end

return data
