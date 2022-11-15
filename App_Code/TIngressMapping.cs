using System.Collections;

public class TIngressMapping
{
    public class TMappingElement
    {
        public int method;
        public string targetname;
        public string source;
        public string key1;
        public string key2;
        public string key3;
        public string key4;
        public string key5;
        public string sourcecolumn;
        public int position;
    }
    public ArrayList mapping;

    public TIngressMapping(string aProjectID)
    {
        mapping = new ArrayList();
        SqlDB dataReader = new SqlDB("Select method, targetname, source, sourcekey1, sourcekey2, sourcekey3, sourcekey4, sourcekey5, sourcecolumn, position FROM orgmanager_ingressmapping ORDER BY position", aProjectID);
        while (dataReader.read())
        {
            TMappingElement tempMapping = new TMappingElement();
            tempMapping.method = dataReader.getInt32(0);
            tempMapping.targetname = dataReader.getString(1);
            tempMapping.source = dataReader.getString(2);
            tempMapping.key1 = dataReader.getString(3);
            tempMapping.key2 = dataReader.getString(4);
            tempMapping.key3 = dataReader.getString(5);
            tempMapping.key4 = dataReader.getString(6);
            tempMapping.key5 = dataReader.getString(7);
            tempMapping.sourcecolumn = dataReader.getString(8);
            tempMapping.position = dataReader.getInt32(9);
            mapping.Add(tempMapping);

        }
        dataReader.close();
    }
}

