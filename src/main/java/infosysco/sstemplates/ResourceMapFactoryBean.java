package infosysco.sstemplates;

import java.io.IOException;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import org.springframework.beans.factory.config.PropertiesFactoryBean;
import org.springframework.context.ResourceLoaderAware;
import org.springframework.core.io.DefaultResourceLoader;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;

/**
 * @author Dreamore
 * springframework перестал его поддерживать в версии 2.5, вернее стали делать это по другому.
 */
public class ResourceMapFactoryBean extends PropertiesFactoryBean implements ResourceLoaderAware {

    private String resourceBasePath = "";

    private ResourceLoader resourceLoader = new DefaultResourceLoader();


    /**
     * Set a base path to prepend to each resource location value
     * in the properties file.
     * <p>E.g.: resourceBasePath="/images", value="/test.gif"
     * -> location="/images/test.gif"
     */
    public void setResourceBasePath(String resourceBasePath) {
        this.resourceBasePath = (resourceBasePath != null ? resourceBasePath : "");
    }

    public void setResourceLoader(ResourceLoader resourceLoader) {
        this.resourceLoader = (resourceLoader != null ? resourceLoader : new DefaultResourceLoader());
    }


    public Class getObjectType() {
        return Map.class;
    }

    /**
     * Create the Map instance, populated with keys and Resource values.
     */
    protected Object createInstance() throws IOException {
        Map resourceMap = new HashMap();
        Properties props = mergeProperties();
        for (Enumeration en = props.propertyNames(); en.hasMoreElements();) {
            String key = (String) en.nextElement();
            String location = props.getProperty(key);
            resourceMap.put(key, getResource(location));
        }
        return resourceMap;
    }

    /**
     * Fetch the Resource handle for the given location,
     * prepeding the resource base path.
     * @param location the resource location
     * @return the Resource handle
     * @see org.springframework.core.io.ResourceLoader#getResource(String)
     */
    protected Resource getResource(String location) {
        return this.resourceLoader.getResource(this.resourceBasePath + location);
    }
}