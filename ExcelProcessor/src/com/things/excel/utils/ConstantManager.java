package com.things.excel.utils;

import java.text.MessageFormat;
import java.util.*;

/**
 * An internationalization / localization helper class which reduces
 * the bother of handling ResourceBundles and takes care of the
 * common cases of message formating which otherwise require the
 * creation of Object arrays and such.
 *
 * <p>The ConstantManager operates on a package basis. One ConstantManager
 * per package can be created and accessed via the getManager method
 * call.
 *
 * <p>The ConstantManager will look for a ResourceBundle named by
 * the package name given plus the suffix of "LocalStrings". In
 * practice, this means that the localized information will be contained
 * in a LocalStrings.properties file located in the package
 * directory of the classpath.
 *
 * <p>Please see the documentation for java.util.ResourceBundle for
 * more information.
 *
 */

public final class ConstantManager {

    /**
     * The ResourceBundle for this ConstantManager.
     */

    private ResourceBundle bundle;

    /**
     * Creates a new ConstantManager for a given package. This is a
     * private method and all access to it is arbitrated by the
     * static getManager method call so that only one ConstantManager
     * per package will be created.
     *
     * @param fileName 		Path of Properties file to create StringManager for.
     */

    private ConstantManager(String fileName) {
        String bundleName = fileName;
        bundle = ResourceBundle.getBundle(bundleName);
    }

	/**
     * Creates a new ConstantManager for a given package. This is a
     * private method and all access to it is arbitrated by the
     * static getManager method call so that only one ConstantManager
     * per package will be created.
     *
     *
     * @param fileName		This Paramter Specifies the package name of the properties file
     *						to be used for the perticular package.
     * @param loc			This parameter is used to specify the localized properties file
     *						for a perticular locale.
     */

    private ConstantManager(String fileName,Locale loc) {
        String bundleName = fileName;
        bundle = ResourceBundle.getBundle(bundleName,loc);
    }

    /**
     * Get a string from the underlying resource bundle.
     *
     * @param key  		The key used in the Properties file to specify the String Constant
     */

    public String getString(String key) {
        if (key == null) {
            String msg = "key is null";

            throw new NullPointerException(msg);
        }

        String str = null;

        try {
	    str = bundle.getString(key);
        } catch (MissingResourceException mre) {
            //str = "[cannot find message associated with key '" + key + "' due to " + mre + "]";
	    // mre. print Stack Trace();
			str = null;
        }

        return str;
    }

    /**
     * Get a string from the underlying resource bundle and format
     * it with the given set of arguments.
     *
     * @param key		The key used to specify the String Constant in the properties file
     * @param args		An array of Objects used for formating the String constant in the properties file
     */

    public String getString(String key, Object[] args) {

		String iString = null;
		String value = getString(key);

		// this check for the runtime exception is some pre 1.1.6
		// VM's don't do an automatic toString() on the passed in
		// objects and barf out

		try {
			// ensure the arguments are not null so pre 1.2 VM's don't barf
			Object nonNullArgs[] = args;
			if(args != null) {
				for (int i=0; i<args.length; i++) {
					if (args[i] == null) {
						if (nonNullArgs==args)
							nonNullArgs=(Object[])args.clone();
						nonNullArgs[i] = "null";
					}
				}
			}
			else {
				System.out.println("Args is null");
				nonNullArgs = new Object[] {"null"};
			}

			iString = MessageFormat.format(value, nonNullArgs);

		} catch (IllegalArgumentException iae) {

			StringBuffer buf = new StringBuffer();
		    buf.append(value);

			for (int i = 0; i < args.length; i++) {

				buf.append(" arg[" + i + "]=" + args[i]);

			}

			iString = buf.toString();
		}

		return iString;

	}

    /**
     * Get a string from the underlying resource bundle and format it
     * with the given object argument. This argument can of course be
     * a String object.
     *
     * @param key		The key used to specify the String Constant in the properties file
     * @param arg		An object parameter used to format the String constant note that it is not a array
     */

    public String getString(String key, Object arg) {
	Object[] args = new Object[] {arg};
	return getString(key, args);
    }

    /**
     * Get a string from the underlying resource bundle and format it
     * with the given object arguments. These arguments can of course
     * be String objects.
     *
     * @param key		The key used to specify the String Constant in the properties file
     * @param arg1		An object parameter used to format the String constant This will
     *					be used to plug in the frist varibale.
     * @param arg2		An object parameter used to format the String constant This will
     *					be used to plug in the Second varibale.
     */

    public String getString(String key, Object arg1, Object arg2) {
		Object[] args = new Object[] {arg1, arg2};
		return getString(key, args);

	}

    /**
     * Get a string from the underlying resource bundle and format it
     * with the given object arguments. These arguments can of course
     * be String objects.
     *
     * @param key		The key used to specify the String Constant in the properties file
     * @param arg1		An object parameter used to format the String constant This will
     *					be used to plug in the frist varibale.
     * @param arg2		An object parameter used to format the String constant This will
     *					be used to plug in the Second varibale.
     * @param arg3		An object parameter used to format the String constant This will
     *					be used to plug in the third varibale.
     */

    public String getString(String key, Object arg1, Object arg2,
			    Object arg3) {
	Object[] args = new Object[] {arg1, arg2, arg3};
	return getString(key, args);
    }

    /**
     * Get a string from the underlying resource bundle and format it
     * with the given object arguments. These arguments can of course
     * be String objects.
     *
     * @param key		The key used to specify the String Constant in the properties file
     * @param arg1		An object parameter used to format the String constant This will
     *					be used to plug in the frist varibale.
     * @param arg2		An object parameter used to format the String constant This will
     *					be used to plug in the Second varibale.
     * @param arg3		An object parameter used to format the String constant This will
     *					be used to plug in the third varibale.
     * @param arg4		An object parameter used to format the String constant This will
     *					be used to plug in the fourth varibale.
     */

    public String getString(String key, Object arg1, Object arg2,
			    Object arg3, Object arg4) {
	Object[] args = new Object[] {arg1, arg2, arg3, arg4};
	return getString(key, args);
    }
    // --------------------------------------------------------------
    // STATIC SUPPORT METHODS
    // --------------------------------------------------------------

    @SuppressWarnings("rawtypes")
	private static Hashtable managers = new Hashtable();

    /**
     * Get the ConstantManager for a particular package. If a manager for
     * a package already exists, it will be reused, else a new
     * StringManager will be created and returned.
     *
     * @param fileName 		This Paramter Specifies the package name of the properties file
     *						to be used for the perticular package.
     */

    @SuppressWarnings("unchecked")
	public synchronized static ConstantManager getManager(String fileName) {

		ConstantManager mgr = (ConstantManager)managers.get(fileName);

		if (mgr == null) {

			mgr = new ConstantManager(fileName);
			managers.put(fileName, mgr);
		}

		return mgr;

	}

	/**
     * Get the ConstantManager for a particular package and Locale. If a manager for
     * a package already exists, it will be reused, else a new
     * ConstantManager will be created for that Locale and returned.
     *
     *
     * @param fileName		This Paramter Specifies the package name of the properties file
     *						to be used for the perticular package.
     * @param loc			This parameter is used to specify the localized properties file
     *						for a perticular locale.
     */

   @SuppressWarnings("unchecked")
public synchronized static ConstantManager getManager(String fileName,Locale loc) {

	   ConstantManager mgr = (ConstantManager)managers.get(fileName+"_"+loc.toString());

	   if (mgr == null) {

		   mgr = new ConstantManager(fileName,loc);
		   managers.put(fileName+"_"+loc.toString(), mgr);
	   }

	   return mgr;
   }

}
