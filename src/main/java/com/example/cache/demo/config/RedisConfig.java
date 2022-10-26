package com.example.cache.demo.config;

import org.springframework.beans.factory.annotation.Value;

//@EnableCaching
//@Configuration
public class RedisConfig {

    @Value("${spring.redis.host}")
    private String redisHostName;

    @Value("${spring.redis.port}")
    private int redisPort;

//    @Bean
//    JedisConnectionFactory jedisConnectionFactory() {
//        RedisStandaloneConfiguration redisStandaloneConfiguration = new RedisStandaloneConfiguration(redisHostName, redisPort);
//        return new JedisConnectionFactory(redisStandaloneConfiguration);
//    }
//
//    @Bean(value = "redisTemplate")
//    public RedisTemplate<String, Object> redisTemplate(RedisConnectionFactory redisConnectionFactory) {
//        RedisTemplate<String, Object> redisTemplate = new RedisTemplate<>();
//        redisTemplate.setConnectionFactory(redisConnectionFactory);
//        return redisTemplate;
//    }
//
//    @Primary
//    @Bean(name = "cacheManager") // Default cache manager is infinite
//    public CacheManager cacheManager(RedisConnectionFactory redisConnectionFactory) {
//        return RedisCacheManager.builder(redisConnectionFactory).cacheDefaults(RedisCacheConfiguration.defaultCacheConfig()).build();
//    }
//
//    @Bean(name = "cacheManager1Hour")
//    public CacheManager cacheManager1Hour(RedisConnectionFactory redisConnectionFactory) {
//        Duration expiration = Duration.ofHours(1);
//        return RedisCacheManager.builder(redisConnectionFactory)
//                .cacheDefaults(RedisCacheConfiguration.defaultCacheConfig().entryTtl(expiration)).build();
//    }

}
